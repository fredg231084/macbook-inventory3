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
    
    if (!event.body) {
      throw new Error('No file data received');
    }

    // Decode base64 body to buffer
    const buffer = Buffer.from(event.body, 'base64');
    console.log(`📊 Buffer size: ${buffer.length} bytes`);
    
    // Parse Excel file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    console.log(`📋 Found sheets: ${workbook.SheetNames.join(', ')}`);
    
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('No sheets found in Excel file');
    }
    
    // Get the first worksheet
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Get worksheet range info
    const range = worksheet['!ref'];
    console.log(`📏 Worksheet range: ${range}`);
    
    // Try multiple parsing methods with detailed logging
    let jsonData = [];
    let headers = [];
    let parseMethod = '';
    
    try {
      // Method 1: Standard conversion
      console.log('🔄 Trying Method 1: Standard conversion...');
      jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      if (jsonData.length > 0) {
        headers = Object.keys(jsonData[0]);
        parseMethod = 'Method 1: Standard';
        console.log(`✅ Method 1 success: ${jsonData.length} rows, ${headers.length} columns`);
      } else {
        console.log('⚠️ Method 1: No data returned');
      }
    } catch (e) {
      console.log('❌ Method 1 failed:', e.message);
    }
    
    if (jsonData.length === 0) {
      try {
        // Method 2: Array format then convert
        console.log('🔄 Trying Method 2: Array format conversion...');
        const arrayData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        console.log(`📊 Array data rows: ${arrayData.length}`);
        
        if (arrayData.length > 1) {
          // Find header row
          let headerRowIndex = 0;
          for (let i = 0; i < Math.min(5, arrayData.length); i++) {
            const row = arrayData[i];
            const rowStr = row.join('|').toLowerCase();
            console.log(`🔍 Row ${i}: ${rowStr.substring(0, 100)}...`);
            
            if (rowStr.includes('stock') || rowStr.includes('brand') || rowStr.includes('model') || 
                rowStr.includes('category') || rowStr.includes('processor')) {
              headerRowIndex = i;
              console.log(`🎯 Found header row at index: ${i}`);
              break;
            }
          }
          
          headers = arrayData[headerRowIndex].map(h => String(h || '').trim()).filter(h => h !== '');
          console.log(`📋 Headers found: ${headers.join(', ')}`);
          
          // Convert data rows
          jsonData = arrayData.slice(headerRowIndex + 1).map(row => {
            const obj = {};
            headers.forEach((header, index) => {
              obj[header] = String(row[index] || '').trim();
            });
            return obj;
          }).filter(row => {
            // Filter out completely empty rows
            const hasData = Object.values(row).some(val => val && val.trim() !== '');
            return hasData;
          });
          
          parseMethod = 'Method 2: Array conversion';
          console.log(`✅ Method 2 success: ${jsonData.length} data rows`);
        }
      } catch (e) {
        console.log('❌ Method 2 failed:', e.message);
      }
    }

    if (jsonData.length === 0) {
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          error: 'No data found in Excel file',
          debug: {
            sheetNames: workbook.SheetNames,
            worksheetRange: range,
            parseMethod: parseMethod || 'No method succeeded',
            totalAttempts: 2
          }
        })
      };
    }

    console.log(`📈 Final data: ${jsonData.length} rows with headers: ${headers.join(', ')}`);

    // Log first few rows for debugging
    console.log('🔍 First 2 rows:');
    jsonData.slice(0, 2).forEach((row, index) => {
      console.log(`Row ${index + 1}:`, Object.keys(row).slice(0, 5).map(key => `${key}: ${row[key]}`).join(', '));
    });

    // Filter Apple products with detailed logging
    console.log('🍎 Filtering for Apple products...');
    const appleProducts = jsonData.filter((item, index) => {
      try {
        // Find relevant columns with various possible names
        const category = findValue(item, ['Sub-Category', 'Category', 'Type', 'Product Type', 'sub-category', 'category', 'Sub Category']);
        const model = findValue(item, ['Model', 'Product', 'Name', 'Title', 'model', 'product']);
        const brand = findValue(item, ['Brand', 'Manufacturer', 'Make', 'brand', 'manufacturer']);
        const stock = findValue(item, ['Stock', 'Quantity', 'Qty', 'Available', 'stock', 'quantity']);
        
        // Log first few items for debugging
        if (index < 3) {
          console.log(`🔍 Item ${index + 1} - Brand: "${brand}", Model: "${model}", Category: "${category}", Stock: "${stock}"`);
        }
        
        // Skip items with no stock
        if (stock && (stock === '0' || stock.toLowerCase().includes('out'))) {
          return false;
        }
        
        // Check if it's an Apple product
        const brandMatch = brand.toLowerCase().includes('apple');
        const categoryMatch = category.toLowerCase().includes('laptop') || 
                             category.toLowerCase().includes('macbook') || 
                             category.toLowerCase().includes('tablet') || 
                             category.toLowerCase().includes('phone') || 
                             category.toLowerCase().includes('desktop');
        const modelMatch = model.toLowerCase().includes('macbook') ||
                          model.toLowerCase().includes('ipad') ||
                          model.toLowerCase().includes('iphone') ||
                          model.toLowerCase().includes('imac') ||
                          model.toLowerCase().includes('airpod');
        
        const isApple = brandMatch || categoryMatch || modelMatch;
        
        if (index < 5 && isApple) {
          console.log(`✅ Apple product found: ${model} (${category})`);
        }
        
        return isApple;
      } catch (e) {
        console.log(`⚠️ Error filtering item ${index}:`, e.message);
        return false;
      }
    });

    console.log(`🎯 Found ${appleProducts.length} Apple products out of ${jsonData.length} total items`);

    if (appleProducts.length === 0) {
      // Get sample data for debugging
      const sampleCategories = jsonData.slice(0, 10).map(item => {
        return findValue(item, ['Sub-Category', 'Category', 'Type']);
      }).filter(Boolean);
      
      const sampleBrands = jsonData.slice(0, 10).map(item => {
        return findValue(item, ['Brand', 'Manufacturer', 'Make']);
      }).filter(Boolean);

      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          error: 'No Apple products found after filtering',
          debug: {
            totalRows: jsonData.length,
            headers: headers,
            parseMethod: parseMethod,
            sampleCategories: sampleCategories,
            sampleBrands: sampleBrands,
            firstRowSample: jsonData[0] || {}
          }
        })
      };
    }

    // Create product groups
    console.log('🏗️ Creating product groups...');
    const { productGroups, categories, totalItems, groupCount } = createProductGroups(appleProducts);

    console.log(`✅ Created ${groupCount} product groups from ${totalItems} Apple products`);

    const result = {
      success: true,
      totalItems: totalItems,
      groupCount: groupCount,
      categories: categories,
      productGroups: productGroups,
      debug: {
        totalRows: jsonData.length,
        filteredRows: appleProducts.length,
        parseMethod: parseMethod,
        headers: headers.slice(0, 10) // Limit for response size
      }
    };

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(result)
    };

  } catch (error) {
    console.error('❌ Processing error:', error);
    
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ 
        error: `Processing failed: ${error.message}`,
        errorType: error.constructor.name,
        stack: error.stack
      })
    };
  }
};

// Helper function to find value from multiple possible column names
function findValue(item, possibleNames) {
  for (const name of possibleNames) {
    if (item[name] && String(item[name]).trim() !== '') {
      return String(item[name]).trim();
    }
  }
  return '';
}

// Create product groups from Apple products
function createProductGroups(appleProducts) {
  const productGroups = {};
  const categories = {};
  
  appleProducts.forEach((item, index) => {
    try {
      // Extract product information
      const model = findValue(item, ['Model', 'Product', 'Name']) || `Product ${index + 1}`;
      const category = findValue(item, ['Sub-Category', 'Category', 'Type']) || 'Apple Product';
      const processor = findValue(item, ['Processor', 'CPU', 'Chip']) || 'Unknown';
      const storage = findValue(item, ['Storage', 'SSD', 'Hard Drive', 'HDD']) || 'Unknown';
      const memory = findValue(item, ['Memory', 'RAM', 'System Memory']) || 'Unknown';
      const color = findValue(item, ['Color', 'Colour', 'Finish']) || 'Default';
      const condition = findValue(item, ['Condition', 'Grade', 'Quality']) || 'A';
      const serialNumber = findValue(item, ['Serial Number', 'Serial', 'S/N', 'SN']);
      const stock = parseInt(findValue(item, ['Stock', 'Quantity', 'Qty'])) || 1;

      // Determine product type
      const productType = determineProductType(model, category);
      
      // Clean specifications
      const cleanProcessor = cleanSpec(processor);
      const cleanStorage = cleanSpec(storage);
      const cleanMemory = cleanSpec(memory);
      const cleanColor = cleanSpec(color);
      const cleanCondition = cleanCondition(condition);

      // Create grouping key
      const groupKey = `${productType}-${cleanProcessor}-${cleanStorage}`.replace(/[^a-zA-Z0-9]/g, '_');

      // Initialize group if it doesn't exist
      if (!productGroups[groupKey]) {
        productGroups[groupKey] = {
          productType: productType,
          processor: cleanProcessor,
          storage: cleanStorage,
          memory: cleanMemory,
          seoTitle: createSEOTitle(productType, cleanProcessor, cleanStorage, cleanMemory),
          basePrice: estimatePrice(productType, cleanStorage),
          items: [],
          variants: {},
          collections: [productType, 'Apple', 'Refurbished']
        };
      }

      // Add item to group
      productGroups[groupKey].items.push({
        model: model,
        color: cleanColor,
        condition: cleanCondition,
        serialNumber: serialNumber,
        stock: stock,
        originalData: item
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

      productGroups[groupKey].variants[variantKey].quantity += stock;
      if (serialNumber) {
        productGroups[groupKey].variants[variantKey].serialNumbers.push(serialNumber);
      }

      // Count categories
      categories[productType] = (categories[productType] || 0) + 1;

    } catch (itemError) {
      console.log(`⚠️ Error processing item ${index}:`, itemError.message);
    }
  });

  const totalItems = appleProducts.length;
  const groupCount = Object.keys(productGroups).length;

  return { productGroups, categories, totalItems, groupCount };
}

// Determine product type from model and category
function determineProductType(model, category) {
  const modelLower = model.toLowerCase();
  const categoryLower = category.toLowerCase();
  
  if (modelLower.includes('macbook pro')) return 'MacBook Pro';
  if (modelLower.includes('macbook air')) return 'MacBook Air';
  if (modelLower.includes('macbook')) return 'MacBook';
  if (modelLower.includes('ipad pro')) return 'iPad Pro';
  if (modelLower.includes('ipad air')) return 'iPad Air';
  if (modelLower.includes('ipad mini')) return 'iPad Mini';
  if (modelLower.includes('ipad')) return 'iPad';
  if (modelLower.includes('iphone')) return 'iPhone';
  if (modelLower.includes('imac')) return 'iMac';
  if (modelLower.includes('mac mini')) return 'Mac Mini';
  if (modelLower.includes('mac studio')) return 'Mac Studio';
  if (modelLower.includes('airpod')) return 'AirPods';
  
  // Fallback to category
  if (categoryLower.includes('laptop')) return 'MacBook';
  if (categoryLower.includes('tablet')) return 'iPad';
  if (categoryLower.includes('phone')) return 'iPhone';
  if (categoryLower.includes('desktop')) return 'iMac';
  
  return category || 'Apple Product';
}

// Clean specifications
function cleanSpec(spec) {
  if (!spec || spec === 'Unknown') return 'Unknown';
  return spec.trim();
}

// Clean condition
function cleanCondition(condition) {
  if (!condition) return 'A';
  
  const conditionUpper = condition.toUpperCase().trim();
  
  // Map various formats to standard grades
  if (conditionUpper.includes('A') || conditionUpper.includes('EXCELLENT')) return 'A';
  if (conditionUpper.includes('B') || conditionUpper.includes('VERY GOOD')) return 'B';
  if (conditionUpper.includes('C') || conditionUpper.includes('GOOD')) return 'C';
  if (conditionUpper.includes('D') || conditionUpper.includes('FAIR')) return 'D';
  
  return 'A'; // Default
}

// Create SEO title
function createSEOTitle(productType, processor, storage, memory) {
  let title = productType;
  
  const specs = [];
  if (processor && processor !== 'Unknown') specs.push(processor);
  if (storage && storage !== 'Unknown') specs.push(storage);
  if (memory && memory !== 'Unknown') specs.push(memory);
  
  if (specs.length > 0) {
    title += ` - ${specs.join(', ')}`;
  }
  
  title += ' | Certified Refurbished';
  return title;
}

// Estimate price based on product type and storage
function estimatePrice(productType, storage) {
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
    'AirPods': 179
  };
  
  let basePrice = basePrices[productType] || 999;
  
  // Adjust for storage if specified
  if (storage && storage !== 'Unknown') {
    const storageMatch = storage.match(/(\d+)/);
    if (storageMatch) {
      const storageAmount = parseInt(storageMatch[1]);
      if (storageAmount >= 1000) basePrice *= 1.3; // 1TB+
      else if (storageAmount >= 512) basePrice *= 1.15; // 512GB
    }
  }
  
  return Math.round(basePrice);
}
