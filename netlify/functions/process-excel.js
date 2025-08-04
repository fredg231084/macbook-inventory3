const XLSX = require('xlsx');

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
    console.log('🚀 Processing Excel file...');
    
    // For Netlify, the file data comes as base64 in event.body
    if (!event.body) {
      throw new Error('No file data received');
    }

    // Decode base64 body to buffer
    const buffer = Buffer.from(event.body, 'base64');
    console.log(`📊 Buffer created, size: ${buffer.length} bytes`);
    
    // Parse Excel file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    console.log(`📋 Found ${workbook.SheetNames.length} sheets: ${workbook.SheetNames.join(', ')}`);
    
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('No sheets found in Excel file');
    }
    
    // Get the first worksheet
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Convert to JSON with different options to handle various Excel formats
    let jsonData = XLSX.utils.sheet_to_json(worksheet, { 
      raw: false,
      defval: '',
      header: 1  // Get array of arrays first
    });

    console.log(`📈 Raw data rows: ${jsonData.length}`);

    if (jsonData.length === 0) {
      throw new Error('No data found in Excel file');
    }

    // Find the header row (look for common column names)
    let headerRowIndex = -1;
    let headers = [];
    
    for (let i = 0; i < Math.min(10, jsonData.length); i++) {
      const row = jsonData[i];
      if (Array.isArray(row)) {
        const rowStr = row.join('|').toLowerCase();
        if (rowStr.includes('brand') || rowStr.includes('model') || rowStr.includes('stock') || 
            rowStr.includes('sub-category') || rowStr.includes('processor') || rowStr.includes('condition')) {
          headerRowIndex = i;
          headers = row.map(h => String(h || '').trim());
          break;
        }
      }
    }

    if (headerRowIndex === -1) {
      // Try default conversion if no headers found
      jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: '' });
      headers = Object.keys(jsonData[0] || {});
    } else {
      // Convert with proper headers
      const dataRows = jsonData.slice(headerRowIndex + 1);
      jsonData = dataRows.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
          obj[header] = String(row[index] || '').trim();
        });
        return obj;
      }).filter(row => {
        // Filter out completely empty rows
        return Object.values(row).some(val => val && val.toString().trim() !== '');
      });
    }

    console.log(`📋 Headers found: ${headers.join(', ')}`);
    console.log(`📊 Data rows after processing: ${jsonData.length}`);

    if (jsonData.length === 0) {
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          error: 'No data found in Excel file',
          debug: {
            totalRows: jsonData.length,
            sheetNames: workbook.SheetNames,
            headers: headers,
            sampleData: jsonData.slice(0, 3)
          }
        })
      };
    }

    // DEBUG: Log first few rows
    console.log('🔍 First 2 rows of processed data:');
    jsonData.slice(0, 2).forEach((row, index) => {
      console.log(`Row ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    // Filter for Apple products with multiple column name variations
    const appleProducts = jsonData.filter(item => {
      const category = getColumnValue(item, ['Sub-Category', 'Category', 'Type', 'Product Type']).toLowerCase();
      const model = getColumnValue(item, ['Model', 'Product', 'Name', 'Title']).toLowerCase();
      const brand = getColumnValue(item, ['Brand', 'Manufacturer', 'Make']).toLowerCase();
      const stock = getColumnValue(item, ['Stock', 'Quantity', 'Qty', 'Available']);
      
      // Skip items with no stock or 0 quantity
      if (stock && (stock === '0' || stock.toLowerCase() === 'out of stock')) {
        return false;
      }
      
      return brand.includes('apple') || 
             category.includes('laptop') || 
             category.includes('macbook') || 
             category.includes('tablet') || 
             category.includes('phone') || 
             category.includes('desktop') || 
             category.includes('accessory') ||
             category.includes('accessories') ||
             model.includes('macbook') ||
             model.includes('ipad') ||
             model.includes('iphone') ||
             model.includes('airpod') ||
             model.includes('imac') ||
             model.includes('mac mini') ||
             model.includes('mac studio') ||
             model.includes('apple watch');
    });

    console.log(`✅ Found ${appleProducts.length} Apple products after filtering`);

    if (appleProducts.length === 0) {
      // Return detailed debug info
      const categories = [...new Set(jsonData.map(item => getColumnValue(item, ['Sub-Category', 'Category'])).filter(Boolean))];
      const brands = [...new Set(jsonData.map(item => getColumnValue(item, ['Brand', 'Manufacturer'])).filter(Boolean))];
      
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          error: 'No Apple products found after filtering',
          debug: {
            totalRows: jsonData.length,
            headers: headers,
            categories: categories.slice(0, 10),
            brands: brands.slice(0, 10),
            sampleData: jsonData.slice(0, 3)
          }
        })
      };
    }

    // Process Apple products into optimized groups
    const { productGroups, categories, totalItems, groupCount } = processIntoProductGroups(appleProducts);

    console.log(`🎯 Created ${groupCount} product groups from ${totalItems} items`);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        success: true,
        totalItems,
        groupCount,
        categories,
        productGroups,
        debug: {
          totalRows: jsonData.length,
          filteredRows: appleProducts.length,
          headers: headers
        }
      })
    };

  } catch (error) {
    console.error('❌ Processing error:', error);
    
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ 
        error: 'Error processing Excel file: ' + error.message,
        debug: {
          errorType: error.constructor.name,
          errorMessage: error.message
        }
      })
    };
  }
};

// Helper function to get column value with multiple possible column names
function getColumnValue(item, possibleColumns) {
  for (const col of possibleColumns) {
    if (item[col] !== undefined && item[col] !== null && item[col] !== '') {
      return String(item[col]).trim();
    }
  }
  return '';
}

// Process Apple products into optimized product groups
function processIntoProductGroups(appleProducts) {
  const productGroups = {};
  const categories = {};

  appleProducts.forEach(item => {
    try {
      // Extract product information with multiple column name options
      const brand = getColumnValue(item, ['Brand', 'Manufacturer', 'Make']);
      const model = getColumnValue(item, ['Model', 'Product', 'Name', 'Title']);
      const category = getColumnValue(item, ['Sub-Category', 'Category', 'Type', 'Product Type']);
      const processor = getColumnValue(item, ['Processor', 'CPU', 'Chip']);
      const storage = getColumnValue(item, ['Storage', 'SSD', 'Hard Drive', 'Disk']);
      const memory = getColumnValue(item, ['Memory', 'RAM', 'System Memory']);
      const color = getColumnValue(item, ['Color', 'Colour', 'Finish']);
      const condition = getColumnValue(item, ['Condition', 'Grade', 'Quality']);
      const serialNumber = getColumnValue(item, ['Serial Number', 'Serial', 'S/N', 'SN']);
      const stock = getColumnValue(item, ['Stock', 'Quantity', 'Qty', 'Available']);

      // Skip if missing essential data
      if (!model || !category) {
        console.log('⚠️ Skipping item with missing model or category:', { model, category });
        return;
      }

      // Normalize and clean data
      const cleanModel = cleanModelName(model);
      const productType = determineProductType(cleanModel, category);
      const cleanProcessor = cleanProcessor(processor);
      const cleanStorage = cleanStorage(storage);
      const cleanMemory = cleanMemory(memory);
      const cleanColor = cleanColor(color);
      const cleanCondition = cleanCondition(condition);
      const displaySize = extractDisplaySize(cleanModel);
      const year = extractYear(cleanModel);

      // Create grouping key (products with same specs get grouped)
      const groupKey = createGroupKey(productType, cleanProcessor, cleanStorage, cleanMemory, displaySize);

      // Initialize product group if it doesn't exist
      if (!productGroups[groupKey]) {
        productGroups[groupKey] = {
          productType,
          processor: cleanProcessor,
          storage: cleanStorage,
          memory: cleanMemory,
          displaySize,
          year,
          seoTitle: createSEOTitle(productType, cleanProcessor, cleanStorage, cleanMemory, displaySize, year),
          basePrice: estimateBasePrice(productType, cleanProcessor, cleanStorage, cleanMemory),
          items: [],
          variants: {},
          collections: createCollections(productType, cleanProcessor, year, displaySize)
        };
      }

      // Add item to the group
      productGroups[groupKey].items.push({
        originalData: item,
        model: cleanModel,
        color: cleanColor,
        condition: cleanCondition,
        serialNumber,
        stock: parseInt(stock) || 1
      });

      // Create/update variant
      const variantKey = `${cleanColor}-${cleanCondition}`;
      if (!productGroups[groupKey].variants[variantKey]) {
        productGroups[groupKey].variants[variantKey] = {
          color: cleanColor,
          condition: cleanCondition,
          quantity: 0,
          serialNumbers: []
        };
      }

      productGroups[groupKey].variants[variantKey].quantity += parseInt(stock) || 1;
      if (serialNumber) {
        productGroups[groupKey].variants[variantKey].serialNumbers.push(serialNumber);
      }

      // Track categories
      categories[productType] = (categories[productType] || 0) + 1;

    } catch (itemError) {
      console.error('Error processing item:', itemError, item);
    }
  });

  const totalItems = appleProducts.length;
  const groupCount = Object.keys(productGroups).length;

  return { productGroups, categories, totalItems, groupCount };
}

// Clean and normalize model names
function cleanModelName(model) {
  if (!model) return 'Unknown Model';
  
  return model
    .replace(/\b(Apple|APPLE)\b/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Determine product type from model and category
function determineProductType(model, category) {
  const modelLower = model.toLowerCase();
  const categoryLower = category.toLowerCase();
  
  if (modelLower.includes('macbook pro') || categoryLower.includes('macbook pro')) return 'MacBook Pro';
  if (modelLower.includes('macbook air') || categoryLower.includes('macbook air')) return 'MacBook Air';
  if (modelLower.includes('macbook')) return 'MacBook';
  if (modelLower.includes('ipad pro')) return 'iPad Pro';
  if (modelLower.includes('ipad air')) return 'iPad Air';
  if (modelLower.includes('ipad mini')) return 'iPad Mini';
  if (modelLower.includes('ipad')) return 'iPad';
  if (modelLower.includes('iphone')) return 'iPhone';
  if (modelLower.includes('imac')) return 'iMac';
  if (modelLower.includes('mac studio')) return 'Mac Studio';
  if (modelLower.includes('mac mini')) return 'Mac Mini';
  if (modelLower.includes('airpods')) return 'AirPods';
  if (modelLower.includes('apple watch')) return 'Apple Watch';
  if (categoryLower.includes('accessory') || categoryLower.includes('accessories')) return 'Apple Accessory';
  
  return category || 'Apple Product';
}

// Clean processor names
function cleanProcessor(processor) {
  if (!processor) return 'Unknown';
  
  processor = processor.trim();
  
  // Apple Silicon processors
  if (processor.match(/M[1-3](\s*(Pro|Max|Ultra))?/i)) {
    return processor.replace(/Apple\s*/i, '').trim();
  }
  
  // Intel processors
  if (processor.includes('Intel')) {
    return processor.replace(/Intel\s*/i, '').trim();
  }
  
  return processor;
}

// Clean storage specifications
function cleanStorage(storage) {
  if (!storage) return 'Unknown';
  
  storage = storage.toUpperCase().trim();
  
  // Extract storage amount and add TB/GB if missing
  const match = storage.match(/(\d+)\s*(TB|GB|T|G)?/);
  if (match) {
    const amount = match[1];
    let unit = match[2] || '';
    
    // Standardize unit
    if (unit === 'T') unit = 'TB';
    if (unit === 'G') unit = 'GB';
    if (!unit) {
      // Assume GB for amounts under 4, TB for 4+
      unit = parseInt(amount) >= 4 ? 'TB' : 'GB';
    }
    
    return `${amount}${unit}`;
  }
  
  return storage;
}

// Clean memory specifications
function cleanMemory(memory) {
  if (!memory) return 'Unknown';
  
  memory = memory.toUpperCase().trim();
  
  const match = memory.match(/(\d+)\s*(GB|G)?/);
  if (match) {
    return `${match[1]}GB`;
  }
  
  return memory;
}

// Clean color names
function cleanColor(color) {
  if (!color) return 'Default';
  
  return color
    .split(/[\s,]+/)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ')
    .trim();
}

// Clean condition grades
function cleanCondition(condition) {
  if (!condition) return 'A';
  
  condition = condition.toUpperCase().trim();
  
  // Map various condition formats to standard grades
  const conditionMap = {
    'EXCELLENT': 'A',
    'VERY GOOD': 'B',
    'GOOD': 'C',
    'FAIR': 'D',
    'POOR': 'D',
    'GRADE A': 'A',
    'GRADE B': 'B',
    'GRADE C': 'C',
    'GRADE D': 'D',
    'A': 'A',
    'B': 'B',
    'C': 'C',
    'D': 'D'
  };
  
  return conditionMap[condition] || 'A';
}

// Extract display size from model name
function extractDisplaySize(model) {
  const sizeMatch = model.match(/(\d+(?:\.\d+)?)\s*['""]?/);
  if (sizeMatch) {
    return `${sizeMatch[1]}"`;
  }
  return '';
}

// Extract year from model name
function extractYear(model) {
  const yearMatch = model.match(/(20\d{2})/);
  return yearMatch ? yearMatch[1] : '';
}

// Create grouping key for similar products
function createGroupKey(productType, processor, storage, memory, displaySize) {
  return `${productType}-${processor}-${storage}-${memory}-${displaySize}`.replace(/\s+/g, '_');
}

// Create SEO-optimized product title
function createSEOTitle(productType, processor, storage, memory, displaySize, year) {
  let title = productType;
  
  if (displaySize) {
    title += ` ${displaySize}`;
  }
  
  if (year) {
    title += ` (${year})`;
  }
  
  const specs = [];
  if (processor && processor !== 'Unknown') {
    specs.push(processor);
  }
  if (storage && storage !== 'Unknown') {
    specs.push(storage);
  }
  if (memory && memory !== 'Unknown') {
    specs.push(memory);
  }
  
  if (specs.length > 0) {
    title += ` - ${specs.join(', ')}`;
  }
  
  title += ' | Certified Refurbished';
  
  return title;
}

// Estimate base price based on product specifications
function estimateBasePrice(productType, processor, storage, memory) {
  let basePrice = 999; // Default price
  
  // Base prices by product type
  const basePrices = {
    'MacBook Pro': 1999,
    'MacBook Air': 1299,
    'MacBook': 1199,
    'iPad Pro': 899,
    'iPad Air': 649,
    'iPad': 449,
    'iPad Mini': 599,
    'iPhone': 799,
    'iMac': 1499,
    'Mac Studio': 2199,
    'Mac Mini': 799,
    'AirPods': 179,
    'Apple Watch': 399,
    'Apple Accessory': 99
  };
  
  basePrice = basePrices[productType] || 999;
  
  // Adjust for processor
  if (processor && processor !== 'Unknown') {
    if (processor.includes('M3')) basePrice *= 1.2;
    else if (processor.includes('M2')) basePrice *= 1.1;
    else if (processor.includes('M1')) basePrice *= 1.0;
    else if (processor.includes('Intel')) basePrice *= 0.85;
  }
  
  // Adjust for storage
  if (storage && storage !== 'Unknown') {
    const storageNum = parseInt(storage);
    if (storageNum >= 1000) basePrice *= 1.3; // 1TB+
    else if (storageNum >= 512) basePrice *= 1.15; // 512GB
    else if (storageNum >= 256) basePrice *= 1.0; // 256GB
    else basePrice *= 0.9; // Less than 256GB
  }
  
  return Math.round(basePrice);
}

// Create collections for the product
function createCollections(productType, processor, year, displaySize) {
  const collections = [];
  
  // Main product type collection
  collections.push(productType);
  
  // Processor-based collections
  if (processor && processor !== 'Unknown') {
    if (processor.includes('M1') || processor.includes('M2') || processor.includes('M3')) {
      collections.push('Apple Silicon');
    }
    if (processor.includes('M3')) {
      collections.push('M3 Devices');
    } else if (processor.includes('M2')) {
      collections.push('M2 Devices');
    } else if (processor.includes('M1')) {
      collections.push('M1 Devices');
    }
  }
  
  // Year-based collections
  if (year) {
    collections.push(`${year} Models`);
  }
  
  // Size-based collections for laptops and tablets
  if (displaySize && (productType.includes('MacBook') || productType.includes('iPad'))) {
    const size = parseFloat(displaySize);
    if (size >= 15) {
      collections.push('Large Screen');
    } else if (size >= 13) {
      collections.push('Standard Screen');
    } else {
      collections.push('Compact');
    }
  }
  
  // Category collections
  if (productType.includes('MacBook')) {
    collections.push('Laptops');
  } else if (productType.includes('iPad')) {
    collections.push('Tablets');
  } else if (productType.includes('iPhone')) {
    collections.push('Phones');
  } else if (productType.includes('iMac') || productType.includes('Mac')) {
    collections.push('Desktops');
  }
  
  // General collections
  collections.push('Refurbished');
  collections.push('Apple');
  
  return [...new Set(collections)]; // Remove duplicates
}