const fetch = require('node-fetch');

exports.handler = async (event, context) => {
  // Enable CORS
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };

  // Handle preflight requests
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers,
      body: ''
    };
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers,
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  try {
    const { storeUrl, apiToken, productGroups } = JSON.parse(event.body);

    if (!storeUrl || !apiToken || !productGroups) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ error: 'Missing required data' })
      };
    }

    const baseUrl = `https://${storeUrl}/admin/api/2023-10/`;
    const shopifyHeaders = {
      'Content-Type': 'application/json',
      'X-Shopify-Access-Token': apiToken
    };

    // Test connection first
    const testResponse = await fetch(`${baseUrl}shop.json`, { headers: shopifyHeaders });
    if (!testResponse.ok) {
      throw new Error(`Shopify connection failed: ${testResponse.status} - Invalid store URL or API token`);
    }

    // Get existing products (with pagination)
    const existingProducts = await getAllExistingProducts(baseUrl, shopifyHeaders);
    
    // Get or create collections
    const collections = await setupCollections(baseUrl, shopifyHeaders, productGroups);

    let results = {
      created: 0,
      updated: 0,
      errors: 0,
      details: [],
      collectionsCreated: Object.keys(collections).length
    };

    console.log(`Processing ${Object.keys(productGroups).length} product groups...`);

    // Process each product group with enhanced logic
    for (const [key, productGroup] of Object.entries(productGroups)) {
      try {
        const existingProduct = findExistingProduct(existingProducts, productGroup);

        if (existingProduct) {
          await updateExistingProductEnhanced(baseUrl, shopifyHeaders, existingProduct, productGroup, collections);
          results.updated++;
          results.details.push(`✅ Updated: ${productGroup.seoTitle}`);
        } else {
          await createNewProductEnhanced(baseUrl, shopifyHeaders, productGroup, collections);
          results.created++;
          results.details.push(`🆕 Created: ${productGroup.seoTitle}`);
        }

        // Rate limiting - Shopify allows 2 calls per second
        await new Promise(resolve => setTimeout(resolve, 500));

      } catch (error) {
        results.errors++;
        results.details.push(`❌ Error with ${productGroup.seoTitle}: ${error.message}`);
        console.error('Product sync error:', error);
      }
    }

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(results)
    };

  } catch (error) {
    console.error('Sync error:', error);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: 'Sync failed: ' + error.message })
    };
  }
};

// GET ALL EXISTING PRODUCTS WITH PAGINATION
async function getAllExistingProducts(baseUrl, headers) {
  let allProducts = [];
  let nextPageInfo = null;
  
  do {
    let url = `${baseUrl}products.json?limit=250`;
    if (nextPageInfo) {
      url += `&page_info=${nextPageInfo}`;
    }
    
    const response = await fetch(url, { headers });
    const data = await response.json();
    
    if (data.products) {
      allProducts = allProducts.concat(data.products);
    }
    
    // Check for pagination
    const linkHeader = response.headers.get('Link');
    nextPageInfo = null;
    if (linkHeader && linkHeader.includes('rel="next"')) {
      const nextMatch = linkHeader.match(/<[^>]*[?&]page_info=([^&>]+)[^>]*>;\s*rel="next"/);
      if (nextMatch) {
        nextPageInfo = nextMatch[1];
      }
    }
  } while (nextPageInfo);
  
  console.log(`Found ${allProducts.length} existing products`);
  return allProducts;
}

// SETUP COLLECTIONS - Create collections that don't exist
async function setupCollections(baseUrl, headers, productGroups) {
  console.log('Setting up collections...');
  
  // Get all unique collections needed
  const neededCollections = new Set();
  Object.values(productGroups).forEach(group => {
    if (group.collections) {
      group.collections.forEach(collection => neededCollections.add(collection));
    }
  });

  // Get existing collections
  const existingResponse = await fetch(`${baseUrl}custom_collections.json?limit=250`, { headers });
  const existingData = await existingResponse.json();
  const existingCollections = existingData.custom_collections || [];
  
  const existingCollectionNames = existingCollections.map(c => c.title.toLowerCase());
  const collectionsMap = {};
  
  // Map existing collections
  existingCollections.forEach(collection => {
    collectionsMap[collection.title.toLowerCase()] = collection.id;
  });

  // Create missing collections
  for (const collectionName of neededCollections) {
    if (!existingCollectionNames.includes(collectionName.toLowerCase())) {
      try {
        const collectionData = {
          custom_collection: {
            title: collectionName,
            handle: collectionName.toLowerCase().replace(/[^a-z0-9]/g, '-').replace(/-+/g, '-'),
            published: true,
            sort_order: 'best-selling'
          }
        };
        
        const response = await fetch(`${baseUrl}custom_collections.json`, {
          method: 'POST',
          headers,
          body: JSON.stringify(collectionData)
        });
        
        if (response.ok) {
          const newCollection = await response.json();
          collectionsMap[collectionName.toLowerCase()] = newCollection.custom_collection.id;
          console.log(`Created collection: ${collectionName}`);
        }
        
        // Rate limiting
        await new Promise(resolve => setTimeout(resolve, 300));
      } catch (error) {
        console.error(`Error creating collection ${collectionName}:`, error);
      }
    }
  }
  
  return collectionsMap;
}

// CREATE NEW PRODUCT WITH ENHANCED FEATURES
async function createNewProductEnhanced(baseUrl, headers, productGroup, collections) {
  const variants = createEnhancedVariants(productGroup);
  
  // Create option values from variants
  const colorOptions = [...new Set(variants.map(v => v.option1))];
  const conditionOptions = [...new Set(variants.map(v => v.option2))];

  const productData = {
    product: {
      title: productGroup.seoTitle,
      body_html: createEnhancedProductDescription(productGroup),
      vendor: 'Apple',
      product_type: productGroup.productType,
      status: 'active',
      handle: createSEOHandle(productGroup.seoTitle),
      options: [
        {
          name: 'Color',
          values: colorOptions
        },
        {
          name: 'Condition',
          values: conditionOptions
        }
      ],
      variants: variants,
      tags: createProductTags(productGroup),
      metafields: createProductMetafields(productGroup),
      seo_title: productGroup.seoTitle,
      seo_description: createSEODescription(productGroup)
    }
  };

  const response = await fetch(`${baseUrl}products.json`, {
    method: 'POST',
    headers,
    body: JSON.stringify(productData)
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to create product: ${error}`);
  }

  const createdProduct = await response.json();
  
  // Add product to collections
  await addProductToCollections(baseUrl, headers, createdProduct.product.id, productGroup.collections, collections);
  
  return createdProduct;
}

// CREATE ENHANCED VARIANTS
function createEnhancedVariants(productGroup) {
  const variants = [];
  
  // Create variants from the enhanced variant structure
  Object.values(productGroup.variants || {}).forEach(variant => {
    const { color, condition, quantity } = variant;
    
    variants.push({
      title: `${color} - Grade ${condition}`,
      option1: color,
      option2: `Grade ${condition}`,
      inventory_quantity: quantity,
      inventory_management: 'shopify',
      inventory_policy: 'deny',
      sku: generateEnhancedSKU(productGroup, color, condition),
      price: calculateVariantPrice(productGroup, condition),
      compare_at_price: calculateComparePrice(productGroup, condition),
      weight: estimateWeight(productGroup.productType),
      weight_unit: 'kg',
      requires_shipping: true,
      taxable: true,
      fulfillment_service: 'manual'
    });
  });

  // If no variants exist, create a default one
  if (variants.length === 0) {
    variants.push({
      title: 'Default - Grade A',
      option1: 'Default',
      option2: 'Grade A',
      inventory_quantity: productGroup.items?.length || 0,
      inventory_management: 'shopify',
      inventory_policy: 'deny',
      sku: generateEnhancedSKU(productGroup, 'Default', 'A'),
      price: productGroup.basePrice || 999,
      weight: estimateWeight(productGroup.productType),
      weight_unit: 'kg',
      requires_shipping: true,
      taxable: true
    });
  }

  return variants;
}

// GENERATE ENHANCED SKU
function generateEnhancedSKU(productGroup, color, condition) {
  const type = (productGroup.productType || 'PROD').replace(/\s+/g, '').substring(0, 4).toUpperCase();
  const size = (productGroup.displaySize || '').replace(/[^\d]/g, '').substring(0, 2);
  const proc = (productGroup.processor || 'UNK').replace(/\s+/g, '').substring(0, 3).toUpperCase();
  const storage = (productGroup.storage || '').replace(/[^\d]/g, '');
  const colorCode = color.substring(0, 2).toUpperCase();
  const conditionCode = condition.substring(0, 1).toUpperCase();
  
  return `${type}-${size}${proc}-${storage}-${colorCode}${conditionCode}`;
}

// CALCULATE VARIANT PRICE BASED ON CONDITION
function calculateVariantPrice(productGroup, condition) {
  const basePrice = productGroup.basePrice || 999;
  
  // Adjust price based on condition
  const conditionMultipliers = {
    'A': 1.0,     // Excellent
    'B': 0.92,    // Very Good
    'C': 0.82,    // Good
    'D': 0.70,    // Fair
    'P': 0.60     // Poor
  };
  
  const multiplier = conditionMultipliers[condition] || 0.85;
  return Math.round(basePrice * multiplier);
}

// CALCULATE COMPARE AT PRICE
function calculateComparePrice(productGroup, condition) {
  const basePrice = productGroup.basePrice || 999;
  
  // Set compare at price for better deals appearance
  if (condition === 'A') {
    return Math.round(basePrice * 1.15); // Show 15% savings
  } else if (condition === 'B') {
    return Math.round(basePrice * 1.08); // Show 8% additional savings
  }
  
  return null;
}

// ESTIMATE WEIGHT FOR SHIPPING
function estimateWeight(productType) {
  const weights = {
    'MacBook Pro': 2.0,
    'MacBook Air': 1.3,
    'iPad Pro': 0.7,
    'iPad Air': 0.6,
    'iPad': 0.5,
    'iPad Mini': 0.3,
    'iPhone': 0.2,
    'iMac': 4.5,
    'Mac Studio': 2.7,
    'Mac Mini': 1.2,
    'AirPods': 0.1,
    'Magic Mouse': 0.1,
    'Magic Keyboard': 0.3
  };
  
  return weights[productType] || 1.0;
}

// CREATE PRODUCT TAGS
function createProductTags(productGroup) {
  const tags = ['refurbished', 'apple', 'certified'];
  
  // Add product type tag
  tags.push(productGroup.productType.toLowerCase().replace(/\s+/g, '-'));
  
  // Add processor tags
  if (productGroup.processor && productGroup.processor !== 'Unknown') {
    tags.push(productGroup.processor.toLowerCase().replace(/\s+/g, '-'));
    
    if (productGroup.processor.includes('M1') || productGroup.processor.includes('M2') || productGroup.processor.includes('M3')) {
      tags.push('apple-silicon');
    }
    if (productGroup.processor.includes('Intel')) {
      tags.push('intel');
    }
  }
  
  // Add storage tags
  if (productGroup.storage && productGroup.storage !== 'Unknown') {
    tags.push(productGroup.storage.toLowerCase());
  }
  
  // Add memory tags
  if (productGroup.memory && productGroup.memory !== 'Unknown') {
    tags.push(productGroup.memory.toLowerCase().replace('gb', 'gb-ram'));
  }
  
  // Add year tag
  if (productGroup.year) {
    tags.push(productGroup.year);
  }
  
  // Add size tag
  if (productGroup.displaySize) {
    tags.push(productGroup.displaySize.toLowerCase().replace(/\s+/g, '-'));
  }
  
  return tags.join(', ');
}

// CREATE PRODUCT METAFIELDS FOR ADDITIONAL DATA
function createProductMetafields(productGroup) {
  return [
    {
      namespace: 'custom',
      key: 'processor',
      value: productGroup.processor,
      type: 'single_line_text_field'
    },
    {
      namespace: 'custom',
      key: 'display_size',
      value: productGroup.displaySize,
      type: 'single_line_text_field'
    },
    {
      namespace: 'custom',
      key: 'year',
      value: productGroup.year || 'N/A',
      type: 'single_line_text_field'
    },
    {
      namespace: 'custom',
      key: 'total_units',
      value: productGroup.items?.length.toString() || '0',
      type: 'number_integer'
    }
  ];
}

// CREATE SEO-OPTIMIZED HANDLE
function createSEOHandle(title) {
  return title
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, '') // Remove special chars except spaces and hyphens
    .replace(/\s+/g, '-')         // Replace spaces with hyphens
    .replace(/-+/g, '-')          // Replace multiple hyphens with single
    .replace(/^-|-$/g, '')        // Remove leading/trailing hyphens
    .substring(0, 255);           // Shopify handle limit
}

// CREATE SEO DESCRIPTION
function createSEODescription(productGroup) {
  let description = `${productGroup.seoTitle} - Certified refurbished by MacBookDepot.com. `;
  
  if (productGroup.processor && productGroup.processor !== 'Unknown') {
    description += `Powered by ${productGroup.processor}. `;
  }
  
  if (productGroup.storage && productGroup.storage !== 'Unknown') {
    description += `${productGroup.storage} storage. `;
  }
  
  if (productGroup.memory && productGroup.memory !== 'Unknown') {
    description += `${productGroup.memory} RAM. `;
  }
  
  description += 'Professional refurbishment, tested & guaranteed. Fast shipping across Canada.';
  
  return description.substring(0, 320); // SEO meta description limit
}

// CREATE ENHANCED PRODUCT DESCRIPTION
function createEnhancedProductDescription(productGroup) {
  const { productType, processor, storage, memory, displaySize, year, items } = productGroup;
  
  return `
    <div class="product-description">
      <div class="key-features">
        <h3>✨ Key Features</h3>
        <ul class="specs-list">
          <li><strong>Model:</strong> ${productType}${displaySize ? ` ${displaySize}` : ''}</li>
          ${processor && processor !== 'Unknown' ? `<li><strong>Processor:</strong> ${processor}</li>` : ''}
          ${storage && storage !== 'Unknown' ? `<li><strong>Storage:</strong> ${storage} SSD</li>` : ''}
          ${memory && memory !== 'Unknown' ? `<li><strong>Memory:</strong> ${memory} RAM</li>` : ''}
          ${year ? `<li><strong>Year:</strong> ${year}</li>` : ''}
          <li><strong>Available Units:</strong> ${items?.length || 0}</li>
        </ul>
      </div>
      
      <div class="condition-info">
        <h3>🏆 Condition Grades</h3>
        <div class="condition-grid">
          <div class="condition-item">
            <strong>Grade A (Excellent):</strong> Like new appearance, minimal wear, fully functional
          </div>
          <div class="condition-item">
            <strong>Grade B (Very Good):</strong> Light cosmetic wear, excellent performance
          </div>
          <div class="condition-item">
            <strong>Grade C (Good):</strong> Visible wear but fully functional, great value
          </div>
          <div class="condition-item">
            <strong>Grade D (Fair):</strong> Heavy wear but fully functional, budget-friendly
          </div>
        </div>
      </div>
      
      <div class="warranty-info">
        <h3>🛡️ MacBookDepot Guarantee</h3>
        <ul>
          <li>✅ <strong>Professional Refurbishment:</strong> Each device is thoroughly tested and restored</li>
          <li>✅ <strong>Quality Assurance:</strong> 30-day return policy</li>
          <li>✅ <strong>Authentic Apple Products:</strong> 100% genuine Apple hardware</li>
          <li>✅ <strong>Fast Shipping:</strong> Quick delivery across Canada</li>
          <li>✅ <strong>Customer Support:</strong> Expert assistance when you need it</li>
        </ul>
      </div>
      
      <div class="whats-included">
        <h3>📦 What's Included</h3>
        <ul>
          <li>${productType}</li>
          <li>Original charging cable and adapter</li>
          <li>Professional cleaning and inspection</li>
          <li>MacBookDepot quality guarantee</li>
        </ul>
      </div>
      
      ${productType.includes('MacBook') ? `
      <div class="perfect-for">
        <h3>🎯 Perfect For</h3>
        <ul>
          <li>🎓 Students and educators</li>
          <li>💼 Business professionals</li>
          <li>🎨 Creative professionals</li>
          <li>💻 Developers and programmers</li>
          <li>🏠 Home and office use</li>
        </ul>
      </div>
      ` : ''}
    </div>
    
    <style>
      .product-description { font-family: Arial, sans-serif; line-height: 1.6; }
      .key-features, .condition-info, .warranty-info, .whats-included, .perfect-for { margin: 20px 0; }
      .specs-list, .warranty-info ul, .whats-included ul, .perfect-for ul { list-style: none; padding: 0; }
      .specs-list li, .warranty-info li, .whats-included li, .perfect-for li { 
        padding: 8px 0; border-bottom: 1px solid #eee; 
      }
      .condition-grid { display: grid; gap: 10px; margin: 10px 0; }
      .condition-item { padding: 10px; background: #f8f9fa; border-radius: 5px; }
      h3 { color: #333; margin: 15px 0 10px 0; }
    </style>
  `;
}

// ADD PRODUCT TO COLLECTIONS
async function addProductToCollections(baseUrl, headers, productId, collectionNames, collectionsMap) {
  if (!collectionNames || collectionNames.length === 0) return;
  
  for (const collectionName of collectionNames) {
    const collectionId = collectionsMap[collectionName.toLowerCase()];
    if (collectionId) {
      try {
        const collectData = {
          collect: {
            product_id: productId,
            collection_id: collectionId
          }
        };
        
        await fetch(`${baseUrl}collects.json`, {
          method: 'POST',
          headers,
          body: JSON.stringify(collectData)
        });
        
        // Rate limiting
        await new Promise(resolve => setTimeout(resolve, 200));
      } catch (error) {
        console.error(`Error adding product to collection ${collectionName}:`, error);
      }
    }
  }
}

// UPDATE EXISTING PRODUCT
async function updateExistingProductEnhanced(baseUrl, headers, existingProduct, productGroup, collections) {
  // Update product details
  const updateData = {
    product: {
      id: existingProduct.id,
      title: productGroup.seoTitle,
      body_html: createEnhancedProductDescription(productGroup),
      tags: createProductTags(productGroup),
      seo_title: productGroup.seoTitle,
      seo_description: createSEODescription(productGroup)
    }
  };
  
  await fetch(`${baseUrl}products/${existingProduct.id}.json`, {
    method: 'PUT',
    headers,
    body: JSON.stringify(updateData)
  });
  
  // Update variants inventory
  const newVariants = createEnhancedVariants(productGroup);
  
  for (const newVariant of newVariants) {
    // Find matching existing variant
    const existingVariant = existingProduct.variants.find(v => 
      v.option1 === newVariant.option1 && v.option2 === newVariant.option2
    );
    
    if (existingVariant) {
      // Update existing variant
      const variantUpdateData = {
        variant: {
          id: existingVariant.id,
          inventory_quantity: newVariant.inventory_quantity,
          price: newVariant.price,
          compare_at_price: newVariant.compare_at_price,
          sku: newVariant.sku
        }
      };
      
      await fetch(`${baseUrl}variants/${existingVariant.id}.json`, {
        method: 'PUT',
        headers,
        body: JSON.stringify(variantUpdateData)
      });
    } else {
      // Create new variant
      const variantCreateData = {
        variant: {
          ...newVariant,
          product_id: existingProduct.id
        }
      };
      
      await fetch(`${baseUrl}products/${existingProduct.id}/variants.json`, {
        method: 'POST',
        headers,
        body: JSON.stringify(variantCreateData)
      });
    }
    
    // Rate limiting
    await new Promise(resolve => setTimeout(resolve, 300));
  }
  
  // Update collections
  await addProductToCollections(baseUrl, headers, existingProduct.id, productGroup.collections, collections);
  
  console.log(`Updated product: ${productGroup.seoTitle}`);
}

// FIND EXISTING PRODUCT
function findExistingProduct(existingProducts, productGroup) {
  const searchTitle = productGroup.seoTitle.toLowerCase();
  const productType = productGroup.productType.toLowerCase();
  const processor = (productGroup.processor || '').toLowerCase();
  const storage = (productGroup.storage || '').toLowerCase();
  const memory = (productGroup.memory || '').toLowerCase();
  
  return existingProducts.find(product => {
    const productTitle = product.title.toLowerCase();
    const productTags = (product.tags || '').toLowerCase();
    
    // Match by title similarity
    if (productTitle === searchTitle) return true;
    
    // Match by key components
    const titleMatches = productTitle.includes(productType) &&
                        (processor === 'unknown' || productTitle.includes(processor)) &&
                        (storage === 'unknown' || productTitle.includes(storage));
    
    const tagMatches = productTags.includes(productType.replace(/\s+/g, '-')) &&
                      (processor === 'unknown' || productTags.includes(processor.replace(/\s+/g, '-')));
    
    return titleMatches || tagMatches;
  });
}