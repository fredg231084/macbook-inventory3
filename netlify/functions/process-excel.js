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
    // For Netlify, the file data comes as base64 in event.body
    if (!event.body) {
      throw new Error('No file data received');
    }

    console.log('🔄 Processing Excel file...');
    
    // Decode base64 body to buffer
    const buffer = Buffer.from(event.body, 'base64');
    console.log(`📊 Buffer size: ${buffer.length} bytes`);
    
    // Parse Excel file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('No sheets found in Excel file');
    }
    
    console.log(`📋 Found ${workbook.SheetNames.length} sheets: ${workbook.SheetNames.join(', ')}`);
    
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    console.log(`📈 Total rows extracted: ${jsonData.length}`);
    
    if (!jsonData || jsonData.length === 0) {
      throw new Error('No data found in Excel file');
    }

    // DEBUG: Log first few rows to see the actual data structure
    console.log('🔍 DEBUG: First 3 rows of data:');
    jsonData.slice(0, 3).forEach((row, index) => {
      console.log(`Row ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    // DEBUG: Log all unique column headers
    const allHeaders = new Set();
    jsonData.forEach(row => {
      Object.keys(row).forEach(key => allHeaders.add(key));
    });
    console.log('🏷️ DEBUG: All column headers found:', Array.from(allHeaders));

    // DEBUG: Log unique values in key columns
    const categories = [...new Set(jsonData.map(item => item['Sub-Category']).filter(Boolean))];
    const brands = [...new Set(jsonData.map(item => item.Brand).filter(Boolean))];
    const models = [...new Set(jsonData.map(item => item.Model).filter(Boolean))].slice(0, 10); // First 10 models
    
    console.log('🏪 DEBUG: Unique Sub-Categories:', categories);
    console.log('🍎 DEBUG: Unique Brands:', brands);
    console.log('📱 DEBUG: Sample Models (first 10):', models);

    // ENHANCED FILTERING with debug logging
    console.log('🔍 Starting filtering process...');
    
    const allProducts = jsonData.filter(item => {
      const category = (item['Sub-Category'] || '').toString().toLowerCase();
      const model = (item.Model || '').toString().toLowerCase();
      const brand = (item.Brand || '').toString().toLowerCase();
      
      // Log filtering details for first few items
      const itemIndex = jsonData.indexOf(item);
      if (itemIndex < 5) {
        console.log(`DEBUG Item ${itemIndex + 1}:`);
        console.log(`  Brand: "${brand}"`);
        console.log(`  Category: "${category}"`);
        console.log(`  Model: "${model}"`);
      }
      
      // Enhanced filtering conditions
      const conditions = [
        brand.includes('apple'),
        category.includes('laptop'),
        category.includes('macbook'),
        category.includes('tablet'),
        category.includes('phone'),
        category.includes('desktop'),
        category.includes('accessory'),
        category.includes('accessories'), // Added plural
        model.includes('macbook'),
        model.includes('ipad'),
        model.includes('iphone'),
        model.includes('airpod'),
        model.includes('imac'),
        model.includes('magic'),
        // Add more flexible matching
        category.includes('mini'),
        category.includes('air'),
        category.includes('pro')
      ];
      
      const isMatch = conditions.some(condition => condition);
      
      if (itemIndex < 5) {
        console.log(`  Match found: ${isMatch}`);
        console.log(`  Matching conditions: ${conditions.map((c, i) => c ? i : null).filter(x => x !== null)}`);
      }
      
      return isMatch;
    });

    console.log(`✅ Found ${allProducts.length} Apple products after filtering`);

    // If no products found, let's try even broader filtering
    if (allProducts.length === 0) {
      console.log('⚠️ No products found with standard filtering, trying broader approach...');
      
      // Try to find ANY products that might be Apple products
      const broadProducts = jsonData.filter(item => {
        const itemStr = JSON.stringify(item).toLowerCase();
        return itemStr.includes('apple') || 
               itemStr.includes('macbook') || 
               itemStr.includes('ipad') || 
               itemStr.includes('iphone') ||
               itemStr.includes('mac ') ||
               itemStr.includes('airpod') ||
               itemStr.includes('a24') || // Common Apple model codes
               itemStr.includes('a23') ||
               itemStr.includes('m1') ||
               itemStr.includes('m2') ||
               itemStr.includes('m3');
      });
      
      console.log(`🔍 Broader search found ${broadProducts.length} potential Apple products`);
      
      if (broadProducts.length > 0) {
        console.log('📝 Sample broad matches:');
        broadProducts.slice(0, 3).forEach((item, index) => {
          console.log(`Broad match ${index + 1}:`, JSON.stringify(item, null, 2));
        });
      }
      
      // Return error with debug info
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          error: 'No Apple products found after filtering',
          debug: {
            totalRows: jsonData.length,
            headers: Array.from(allHeaders),
            categories: categories,
            brands: brands,
            sampleModels: models,
            broadMatches: broadProducts.length,
            sampleData: jsonData.slice(0, 2) // Include sample data for analysis
          }
        })
      };
    }

    // Continue with normal processing if products found
    const productGroups = groupProductsEnhanced(allProducts);

    const processedData = {
      totalItems: allProducts.length,
      productGroups: productGroups,
      rawData: allProducts,
      groupCount: Object.keys(productGroups).length,
      categories: getCategoryStats(allProducts)
    };

    console.log(`🎉 Created ${Object.keys(productGroups).length} product groups`);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(processedData)
    };

  } catch (error) {
    console.error('❌ Error processing Excel:', error);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ 
        error: 'Error processing Excel file: ' + error.message,
        stack: error.stack
      })
    };
  }
};

// Keep all the existing functions from the enhanced version
function groupProductsEnhanced(products) {
  const productGroups = {};

  products.forEach(item => {
    const productInfo = analyzeProduct(item);
    const productKey = createEnhancedProductKey(productInfo);

    if (!productGroups[productKey]) {
      productGroups[productKey] = {
        productType: productInfo.productType,
        model: productInfo.model,
        displaySize: productInfo.displaySize,
        processor: productInfo.processor,
        storage: productInfo.storage,
        memory: productInfo.memory,
        year: productInfo.year,
        seoTitle: createSEOTitle(productInfo),
        collections: determineCollections(productInfo),
        basePrice: estimatePrice(productInfo, item),
        items: [],
        variants: {},
        originalCategory: item['Sub-Category']
      };
    }

    const variantKey = `${productInfo.color || 'Default'}_${productInfo.condition || 'A'}`;
    
    if (!productGroups[productKey].variants[variantKey]) {
      productGroups[productKey].variants[variantKey] = {
        color: productInfo.color || 'Default',
        condition: productInfo.condition || 'A',
        quantity: 0,
        items: [],
        serialNumbers: []
      };
    }

    productGroups[productKey].variants[variantKey].quantity++;
    productGroups[productKey].variants[variantKey].items.push(item);
    productGroups[productKey].variants[variantKey].serialNumbers.push(item['Serial Number']);
    productGroups[productKey].items.push(item);
  });

  return productGroups;
}

function analyzeProduct(item) {
  const model = item.Model || '';
  const processor = item.Processor || '';
  const category = (item['Sub-Category'] || '').toLowerCase();
  const storage = item.Storage || '';
  const memory = item.Memory || '';
  const color = item.Color || '';
  const condition = item.Condition || 'A';

  let productType = 'Unknown';
  let displaySize = '';
  let year = '';

  if (category.includes('laptop') || category.includes('macbook') || 
      model.toLowerCase().includes('macbook') || processor.toLowerCase().includes('macbook')) {
    
    if (processor.toLowerCase().includes('air') || category.includes('air') || model.toLowerCase().includes('air')) {
      productType = 'MacBook Air';
      displaySize = '13-inch';
    } else {
      productType = 'MacBook Pro';
      if (processor.includes('16"') || processor.includes('16-inch')) {
        displaySize = '16-inch';
      } else if (processor.includes('14"') || processor.includes('14-inch')) {
        displaySize = '14-inch';
      } else {
        displaySize = '13-inch';
      }
    }
    
    if (processor.includes('2023')) year = '2023';
    else if (processor.includes('2022')) year = '2022';
    else if (processor.includes('2021')) year = '2021';
    else if (processor.includes('2020')) year = '2020';
    else if (processor.includes('2019')) year = '2019';
  }
  else if (category.includes('tablet') || model.toLowerCase().includes('ipad')) {
    if (model.toLowerCase().includes('pro')) {
      productType = 'iPad Pro';
      if (model.includes('11')) displaySize = '11-inch';
      else if (model.includes('12.9')) displaySize = '12.9-inch';
    } else if (model.toLowerCase().includes('air')) {
      productType = 'iPad Air';
      displaySize = '10.9-inch';
    } else if (model.toLowerCase().includes('mini')) {
      productType = 'iPad Mini';
      displaySize = '8.3-inch';
    } else {
      productType = 'iPad';
      displaySize = '10.2-inch';
    }
  }
  else if (category.includes('phone') || model.toLowerCase().includes('iphone')) {
    productType = 'iPhone';
    if (model.toLowerCase().includes('pro max')) {
      displaySize = 'Pro Max';
    } else if (model.toLowerCase().includes('pro')) {
      displaySize = 'Pro';
    } else if (model.toLowerCase().includes('plus')) {
      displaySize = 'Plus';
    } else if (model.toLowerCase().includes('mini')) {
      displaySize = 'mini';
    }
  }
  else if (category.includes('desktop') || model.toLowerCase().includes('imac')) {
    productType = 'iMac';
    if (model.includes('24') || processor.includes('24')) displaySize = '24-inch';
    else if (model.includes('27') || processor.includes('27')) displaySize = '27-inch';
  }
  else if (model.toLowerCase().includes('mac studio')) {
    productType = 'Mac Studio';
  } else if (model.toLowerCase().includes('mac mini') || category.includes('mini')) {
    productType = 'Mac Mini';
  }
  else if (category.includes('accessory') || model.toLowerCase().includes('airpod') || 
           model.toLowerCase().includes('magic') || model.toLowerCase().includes('keyboard')) {
    if (model.toLowerCase().includes('airpod')) {
      productType = 'AirPods';
    } else if (model.toLowerCase().includes('magic mouse')) {
      productType = 'Magic Mouse';
    } else if (model.toLowerCase().includes('magic keyboard') || model.toLowerCase().includes('kybd')) {
      productType = 'Magic Keyboard';
    } else {
      productType = 'Apple Accessory';
    }
  }

  return {
    productType,
    model: normalizeModel(model),
    displaySize,
    processor: normalizeProcessor(processor),
    storage: normalizeStorage(storage),
    memory: normalizeMemory(memory),
    color: normalizeColor(color),
    condition: normalizeCondition(condition),
    year,
    originalModel: model
  };
}

function createEnhancedProductKey(productInfo) {
  return `${productInfo.productType}_${productInfo.displaySize}_${productInfo.processor}_${productInfo.storage}_${productInfo.memory}`.replace(/\s+/g, '_');
}

function createSEOTitle(productInfo) {
  let title = productInfo.productType;
  
  if (productInfo.displaySize) {
    title += ` ${productInfo.displaySize}`;
  }
  
  if (productInfo.processor && productInfo.processor !== 'Unknown') {
    title += ` ${productInfo.processor}`;
  }
  
  if (productInfo.storage && productInfo.storage !== 'Unknown') {
    title += ` - ${productInfo.storage}`;
  }
  
  if (productInfo.memory && productInfo.memory !== 'Unknown') {
    title += `, ${productInfo.memory} RAM`;
  }
  
  if (productInfo.year) {
    title += ` (${productInfo.year})`;
  }
  
  return title.replace(/\s+/g, ' ').trim();
}

function determineCollections(productInfo) {
  const collections = ['All Products'];
  collections.push(productInfo.productType);
  
  if (productInfo.productType.includes('MacBook')) {
    collections.push('MacBooks');
    if (productInfo.processor.includes('M1') || productInfo.processor.includes('M2') || productInfo.processor.includes('M3')) {
      collections.push('Apple Silicon');
    }
    if (productInfo.processor.includes('Intel')) {
      collections.push('Intel MacBooks');
    }
  }
  
  if (productInfo.productType.includes('iPad')) {
    collections.push('iPads');
  }
  
  if (productInfo.productType.includes('iPhone')) {
    collections.push('iPhones');
  }
  
  if (productInfo.productType.includes('iMac') || productInfo.productType.includes('Mac Studio') || productInfo.productType.includes('Mac Mini')) {
    collections.push('Desktops');
  }
  
  if (productInfo.productType.includes('AirPods') || productInfo.productType.includes('Magic') || productInfo.productType.includes('Apple Accessory')) {
    collections.push('Accessories');
  }
  
  if (productInfo.displaySize) {
    collections.push(`${productInfo.displaySize} Devices`);
  }
  
  if (productInfo.processor && productInfo.processor !== 'Unknown') {
    collections.push(productInfo.processor);
  }
  
  if (productInfo.year) {
    collections.push(`${productInfo.year} Models`);
  }
  
  return [...new Set(collections)];
}

function normalizeModel(model) {
  if (!model) return 'Unknown';
  
  const modelStr = model.toString().toLowerCase();
  
  if (modelStr.includes('macbook pro')) return 'MacBook Pro';
  if (modelStr.includes('macbook air')) return 'MacBook Air';
  if (modelStr.includes('macbook')) return 'MacBook';
  if (modelStr.includes('ipad pro')) return 'iPad Pro';
  if (modelStr.includes('ipad air')) return 'iPad Air';
  if (modelStr.includes('ipad mini')) return 'iPad Mini';
  if (modelStr.includes('ipad')) return 'iPad';
  if (modelStr.includes('iphone')) return 'iPhone';
  if (modelStr.includes('imac')) return 'iMac';
  if (modelStr.includes('mac studio')) return 'Mac Studio';
  if (modelStr.includes('mac mini')) return 'Mac Mini';
  if (modelStr.includes('airpod')) return 'AirPods';
  if (modelStr.includes('magic mouse')) return 'Magic Mouse';
  if (modelStr.includes('magic') && modelStr.includes('keyboard')) return 'Magic Keyboard';
  
  return model;
}

function normalizeProcessor(processor) {
  if (!processor) return 'Unknown';
  
  const procStr = processor.toString().toLowerCase();
  
  if (procStr.includes('m3 pro')) return 'M3 Pro';
  if (procStr.includes('m3 max')) return 'M3 Max';
  if (procStr.includes('m3')) return 'M3';
  if (procStr.includes('m2 pro')) return 'M2 Pro';
  if (procStr.includes('m2 max')) return 'M2 Max';
  if (procStr.includes('m2')) return 'M2';
  if (procStr.includes('m1 pro')) return 'M1 Pro';
  if (procStr.includes('m1 max')) return 'M1 Max';
  if (procStr.includes('m1')) return 'M1';
  
  if (procStr.includes('i9')) return 'Intel i9';
  if (procStr.includes('i7')) return 'Intel i7';
  if (procStr.includes('i5')) return 'Intel i5';
  if (procStr.includes('i3')) return 'Intel i3';
  if (procStr.includes('intel')) return 'Intel';
  
  if (procStr.includes('airpods') && procStr.includes('2nd')) return 'AirPods 2nd Gen';
  if (procStr.includes('airpods') && procStr.includes('3rd')) return 'AirPods 3rd Gen';
  if (procStr.includes('airpods') && procStr.includes('pro')) return 'AirPods Pro';
  
  return processor.substring(0, 20);
}

function normalizeStorage(storage) {
  if (!storage) return 'Unknown';
  
  const storageStr = storage.toString().toLowerCase().replace(/\s+/g, '');
  
  if (storageStr.includes('8tb')) return '8TB';
  if (storageStr.includes('4tb')) return '4TB';
  if (storageStr.includes('2tb')) return '2TB';
  if (storageStr.includes('1tb') || storageStr.includes('1000gb')) return '1TB';
  if (storageStr.includes('512gb')) return '512GB';
  if (storageStr.includes('256gb')) return '256GB';
  if (storageStr.includes('128gb')) return '128GB';
  if (storageStr.includes('64gb')) return '64GB';
  if (storageStr.includes('32gb')) return '32GB';
  
  return storage;
}

function normalizeMemory(memory) {
  if (!memory) return 'Unknown';
  
  const memStr = memory.toString().toLowerCase().replace(/\s+/g, '');
  
  if (memStr.includes('128gb')) return '128GB';
  if (memStr.includes('64gb')) return '64GB';
  if (memStr.includes('32gb')) return '32GB';
  if (memStr.includes('16gb')) return '16GB';
  if (memStr.includes('8gb')) return '8GB';
  if (memStr.includes('4gb')) return '4GB';
  
  return memory;
}

function normalizeColor(color) {
  if (!color) return 'Default';
  
  const colorStr = color.toString().toLowerCase();
  
  if (colorStr.includes('space gray') || colorStr.includes('space grey')) return 'Space Gray';
  if (colorStr.includes('silver')) return 'Silver';
  if (colorStr.includes('gold')) return 'Gold';
  if (colorStr.includes('rose gold')) return 'Rose Gold';
  if (colorStr.includes('midnight')) return 'Midnight';
  if (colorStr.includes('starlight')) return 'Starlight';
  if (colorStr.includes('blue')) return 'Blue';
  if (colorStr.includes('purple')) return 'Purple';
  if (colorStr.includes('pink')) return 'Pink';
  if (colorStr.includes('green')) return 'Green';
  if (colorStr.includes('red')) return 'Red';
  if (colorStr.includes('black')) return 'Black';
  if (colorStr.includes('white')) return 'White';
  
  return color;
}

function normalizeCondition(condition) {
  if (!condition) return 'A';
  
  const condStr = condition.toString().toUpperCase();
  
  if (condStr === 'A' || condStr === 'EXCELLENT') return 'A';
  if (condStr === 'B' || condStr === 'VERY GOOD') return 'B';
  if (condStr === 'C' || condStr === 'GOOD') return 'C';
  if (condStr === 'D' || condStr === 'FAIR') return 'D';
  if (condStr === 'P' || condStr === 'POOR') return 'P';
  
  return condition;
}

function estimatePrice(productInfo, item) {
  let basePrice = 300;
  
  switch (productInfo.productType) {
    case 'MacBook Pro':
      basePrice = 1200;
      if (productInfo.displaySize === '16-inch') basePrice += 400;
      if (productInfo.displaySize === '14-inch') basePrice += 200;
      break;
    case 'MacBook Air':
      basePrice = 800;
      break;
    case 'iPad Pro':
      basePrice = 600;
      if (productInfo.displaySize === '12.9-inch') basePrice += 200;
      break;
    case 'iPad Air':
      basePrice = 450;
      break;
    case 'iPad':
      basePrice = 250;
      break;
    case 'iPad Mini':
      basePrice = 350;
      break;
    case 'iPhone':
      basePrice = 400;
      if (productInfo.displaySize.includes('Pro')) basePrice += 300;
      break;
    case 'iMac':
      basePrice = 1000;
      if (productInfo.displaySize === '27-inch') basePrice += 500;
      break;
    case 'Mac Studio':
      basePrice = 1500;
      break;
    case 'Mac Mini':
      basePrice = 500;
      break;
    case 'AirPods':
      basePrice = 100;
      if (productInfo.processor.includes('Pro')) basePrice += 50;
      break;
    default:
      basePrice = 200;
  }
  
  if (productInfo.processor.includes('M3')) basePrice += 300;
  else if (productInfo.processor.includes('M2')) basePrice += 200;
  else if (productInfo.processor.includes('M1')) basePrice += 100;
  
  if (productInfo.processor.includes('Pro') || productInfo.processor.includes('Max')) {
    basePrice += 300;
  }
  
  if (productInfo.storage.includes('2TB')) basePrice += 400;
  else if (productInfo.storage.includes('1TB')) basePrice += 200;
  else if (productInfo.storage.includes('512GB')) basePrice += 100;
  
  if (productInfo.memory.includes('64GB')) basePrice += 600;
  else if (productInfo.memory.includes('32GB')) basePrice += 300;
  else if (productInfo.memory.includes('16GB')) basePrice += 150;
  
  if (productInfo.year === '2023') basePrice += 200;
  else if (productInfo.year === '2022') basePrice += 100;
  else if (productInfo.year === '2021') basePrice += 50;
  else if (productInfo.year && parseInt(productInfo.year) < 2020) basePrice -= 100;
  
  return Math.max(basePrice, 50);
}

function getCategoryStats(products) {
  const stats = {};
  products.forEach(item => {
    const category = item['Sub-Category'] || 'Unknown';
    stats[category] = (stats[category] || 0) + 1;
  });
  return stats;
}