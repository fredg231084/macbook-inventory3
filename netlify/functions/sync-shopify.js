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
      throw new Error(`Shopify connection failed: ${testResponse.status}`);
    }

    // Get existing products
    const existingResponse = await fetch(`${baseUrl}products.json?limit=250`, { headers: shopifyHeaders });
    const existingData = await existingResponse.json();
    const existingProducts = existingData.products || [];

    let results = {
      created: 0,
      updated: 0,
      errors: 0,
      details: []
    };

    // Process each product group with improved logic
    for (const [key, productGroup] of Object.entries(productGroups)) {
      try {
        const productTitle = createImprovedProductTitle(productGroup);
        const existingProduct = findExistingProduct(existingProducts, productTitle, productGroup);

        if (existingProduct) {
          await updateExistingProductImproved(baseUrl, shopifyHeaders, existingProduct, productGroup);
          results.updated++;
          results.details.push(`Updated: ${productTitle}`);
        } else {
          await createNewProductImproved(baseUrl, shopifyHeaders, productGroup);
          results.created++;
          results.details.push(`Created: ${productTitle}`);
        }

        // Small delay to avoid rate limits
        await new Promise(resolve => setTimeout(resolve, 200));

      } catch (error) {
        results.errors++;
        results.details.push(`Error with ${productGroup.model}: ${error.message}`);
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

// IMPROVED PRODUCT CREATION LOGIC
function createImprovedProductTitle(productGroup) {
  const { model, processor, storage, memory } = productGroup;
  
  let title = model || 'MacBook';
  
  if (processor && processor !== 'Unknown') {
    title += ` ${processor}`;
  }
  
  if (storage && storage !== 'Unknown') {
    title += ` ${storage}`;
  }
  
  if (memory && memory !== 'Unknown') {
    title += ` ${memory}`;
  }
  
  return title.replace(/\s+/g, ' ').trim();
}

async function createNewProductImproved(baseUrl, headers, productGroup) {
  const variants = createImprovedVariants(productGroup);
  const productTitle = createImprovedProductTitle(productGroup);

  // Create option values from variants
  const colorOptions = [...new Set(variants.map(v => v.option1))];
  const conditionOptions = [...new Set(variants.map(v => v.option2))];

  const productData = {
    product: {
      title: productTitle,
      body_html: createImprovedProductDescription(productGroup),
      vendor: 'Apple',
      product_type: 'Laptop',
      status: 'active',
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
      tags: createProductTags(productGroup)
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

  return await response.json();
}

function createImprovedVariants(productGroup) {
  const variants = [];
  
  // Create variants from the improved variant structure
  Object.values(productGroup.variants || {}).forEach(variant => {
    const { color, condition, quantity } = variant;
    
    variants.push({
      title: `${color} - Grade ${condition}`,
      option1: color,
      option2: `Grade ${condition}`,
      inventory_quantity: quantity,
      inventory_management: 'shopify',
      inventory_policy: 'deny',
      sku: generateSKU(productGroup, color, condition),
      price: calculateVariantPrice(productGroup, condition),
      compare_at_price: calculateComparePrice(productGroup, condition)
    });
  });

  // If no variants exist, create a default one
  if (variants.length === 0) {
    variants.push({
      title: 'Default',
      option1: 'Default',
      option2: 'Grade A',
      inventory_quantity: productGroup.items?.length || 0,
      inventory_management: 'shopify',
      inventory_policy: 'deny',
      sku: generateSKU(productGroup, 'Default', 'A'),
      price: productGroup.basePrice || 999
    });
  }

  return variants;
}

function generateSKU(productGroup, color, condition) {
  const model = (productGroup.model || 'MB').replace(/\s+/g, '').toUpperCase();
  const proc = (productGroup.processor || 'UNK').substring(0, 3).toUpperCase();
  const storage = (productGroup.storage || '').replace(/[^\d]/g, '');
  const colorCode = color.substring(0, 3).toUpperCase();
  const conditionCode = condition.substring(0, 1).toUpperCase();
  
  return `${model}-${proc}-${storage}-${colorCode}-${conditionCode}`;
}

function calculateVariantPrice(productGroup, condition) {
  const basePrice = productGroup.basePrice || 999;
  
  // Adjust price based on condition
  const conditionMultipliers = {
    'A': 1.0,
    'B': 0.85,
    'C': 0.70,
    'Excellent': 1.0,
    'Good': 0.85,
    'Fair': 0.70
  };
  
  const multiplier = conditionMultipliers[condition] || 0.85;
  return Math.round(basePrice * multiplier);
}

function calculateComparePrice(productGroup, condition) {
  const basePrice = productGroup.basePrice || 999;
  
  // Set compare at price higher for Grade A/Excellent condition
  if (condition === 'A' || condition === 'Excellent') {
    return Math.round(basePrice * 1.2);
  }
  
  return null;
}

function createProductTags(productGroup) {
  const tags = ['refurbished', 'macbook'];
  
  if (productGroup.processor) {
    tags.push(productGroup.processor.toLowerCase().replace(/\s+/g, '-'));
  }
  
  if (productGroup.storage) {
    tags.push(productGroup.storage.toLowerCase());
  }
  
  if (productGroup.memory) {
    tags.push(productGroup.memory.toLowerCase());
  }
  
  return tags.join(', ');
}

function createImprovedProductDescription(productGroup) {
  const { model, processor, storage, memory, items } = productGroup;
  
  return `
    <div class="product-specs">
      <h3>Product Specifications</h3>
      <ul>
        <li><strong>Model:</strong> ${model}</li>
        <li><strong>Processor:</strong> ${processor}</li>
        <li><strong>Storage:</strong> ${storage}</li>
        <li><strong>Memory:</strong> ${memory}</li>
        <li><strong>Available Units:</strong> ${items?.length || 0}</li>
      </ul>
      
      <h3>Condition Information</h3>
      <p>All devices are professionally refurbished and tested. Available in multiple condition grades:</p>
      <ul>
        <li><strong>Grade A:</strong> Excellent condition, minimal wear</li>
        <li><strong>Grade B:</strong> Good condition, light wear</li>
        <li><strong>Grade C:</strong> Fair condition, visible wear but fully functional</li>
      </ul>
      
      <p>Each device includes original charging accessories and comes with a warranty.</p>
    </div>
  `;
}

async function updateExistingProductImproved(baseUrl, headers, existingProduct, productGroup) {
  // Update inventory quantities for existing variants
  const variants = createImprovedVariants(productGroup);
  
  for (const variant of variants) {
    // Find matching existing variant or create new one
    const existingVariant = existingProduct.variants.find(v => 
      v.option1 === variant.option1 && v.option2 === variant.option2
    );
    
    if (existingVariant) {
      // Update existing variant inventory
      const updateData = {
        variant: {
          id: existingVariant.id,
          inventory_quantity: variant.inventory_quantity
        }
      };
      
      await fetch(`${baseUrl}variants/${existingVariant.id}.json`, {
        method: 'PUT',
        headers,
        body: JSON.stringify(updateData)
      });
    }
  }
  
  console.log(`Updated inventory for: ${existingProduct.title}`);
}

function findExistingProduct(existingProducts, title, productGroup) {
  return existingProducts.find(product => {
    const productTitle = product.title.toLowerCase();
    const searchTitle = title.toLowerCase();
    const model = productGroup.model.toLowerCase();
    
    return productTitle.includes(model) || 
           productTitle === searchTitle ||
           searchTitle.includes(productTitle);
  });
}