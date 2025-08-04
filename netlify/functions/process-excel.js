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
    console.log('Processing Excel file...');
    
    if (!event.body) {
      throw new Error('No file data received');
    }

    // Decode base64 body to buffer
    const buffer = Buffer.from(event.body, 'base64');
    console.log(`Buffer size: ${buffer.length} bytes`);
    
    // Parse Excel file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    console.log(`Found sheets: ${workbook.SheetNames.join(', ')}`);
    
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('No sheets found in Excel file');
    }
    
    // Get the first worksheet
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Try multiple parsing methods
    let jsonData = [];
    let headers = [];
    
    try {
      // Method 1: Standard conversion
      jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      if (jsonData.length > 0) {
        headers = Object.keys(jsonData[0]);
        console.log(`Method 1 success: ${jsonData.length} rows`);
      }
    } catch (e) {
      console.log('Method 1 failed, trying method 2');
    }
    
    if (jsonData.length === 0) {
      try {
        // Method 2: Array format then convert
        const arrayData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        if (arrayData.length > 1) {
          headers = arrayData[0].map(h => String(h || '').trim());
          jsonData = arrayData.slice(1).map(row => {
            const obj = {};
            headers.forEach((header, index) => {
              obj[header] = String(row[index] || '').trim();
            });
            return obj;
          }).filter(row => Object.values(row).some(val => val && val.trim() !== ''));
          console.log(`Method 2 success: ${jsonData.length} rows`);
        }
      } catch (e) {
        console.log('Method 2 failed');
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
            worksheetRange: worksheet['!ref'] || 'No range',
            totalAttempts: 2
          }
        })
      };
    }

    console.log(`Headers: ${headers.join(', ')}`);
    console.log(`Total rows: ${jsonData.length}`);

    // Filter Apple products - be more flexible with column names
    const appleProducts = jsonData.filter(item => {
      try {
        // Find category column (try multiple names)
        let category = '';
        ['Sub-Category', 'Category', 'Type', 'Product Type', 'sub-category', 'category'].forEach(col => {
          if (item[col] && !category) category = String(item[col]).toLowerCase();
        });

        // Find model column
        let model = '';
        ['Model', 'Product', 'Name', 'Title', 'model', 'product'].forEach(col => {
          if (item[col] && !model) model = String(item[col]).toLowerCase();
        });

        // Find brand column
        let brand = '';
        ['Brand', 'Manufacturer', 'Make', 'brand', 'manufacturer'].forEach(col => {
          if (item[col] && !brand) brand = String(item[col]).toLowerCase();
        });

        // Check if it's an Apple product
        const isApple = brand.includes('apple') || 
                       category.includes('laptop') || 
                       category.includes('macbook') || 
                       category.includes('tablet') || 
                       category.includes('phone') || 
                       category.includes('desktop') ||
                       model.includes('macbook') ||
                       model.includes('ipad') ||
                       model.includes('iphone') ||
                       model.includes('imac') ||
                       model.includes('airpod');

        return isApple;
      } catch (e) {
        return false;
      }
    });

    console.log(`Found ${appleProducts.length} Apple products`);

    if (appleProducts.length === 0) {
      // Get sample data for debugging
      const sampleCategories = jsonData.slice(0, 5).map(item => {
        let cat = '';
        ['Sub-Category', 'Category', 'Type'].forEach(col => {
          if (item[col] && !cat) cat = String(item[col]);
        });
        return cat;
      }).filter(Boolean);

      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          error: 'No Apple products found',
          debug: {
            totalRows: jsonData.length,
            headers: headers,
            sampleCategories: sampleCategories,
            firstRow: jsonData[0] || {}
          }
        })
      };
    }

    // Create simplified product groups
    const productGroups = {};
    const categories = {};

    appleProducts.forEach((item, index) => {
      try {
        // Get basic product info
        const model = getField(item, ['Model', 'Product', 'Name']) || `Product ${index + 1}`;
        const category = getField(item, ['Sub-Category', 'Category', 'Type']) || 'Apple Product';
        const processor = getField(item, ['Processor', 'CPU', 'Chip']) || 'Unknown';
        const storage = getField(item, ['Storage', 'SSD', 'Hard Drive']) || 'Unknown';
        const memory = getField(item, ['Memory', 'RAM']) || 'Unknown';
        const color = getField(item, ['Color', 'Colour']) || 'Default';
        const condition = getField(item, ['Condition', 'Grade']) || 'A';

        // Create simple grouping key
        const productType = determineType(model, category);
        const groupKey = `${productType}-${processor}-${storage}`.replace(/\s+/g, '_');

        // Initialize group
        if (!productGroups[groupKey]) {
          productGroups[groupKey] = {
            productType: productType,
            processor: processor,
            storage: storage,
            memory: memory,
            seoTitle: createTitle(productType, processor, storage, memory),
            basePrice: 999,
            items: [],
            variants: {},
            collections: [productType, 'Apple', 'Refurbished']
          };
        }

        // Add to group
        productGroups[groupKey].items.push({
          model: model,
          color: color,
          condition: condition,
          originalData: item
        });

        // Create variant
        const variantKey = `${color}-${condition}`;
        if (!productGroups[groupKey].variants[variantKey]) {
          productGroups[groupKey].variants[variantKey] = {
            color: color,
            condition: condition,
            quantity: 0
          };
        }
        productGroups[groupKey].variants[variantKey].quantity += 1;

        // Count categories
        categories[productType] = (categories[productType] || 0) + 1;

      } catch (itemError) {
        console.log(`Error processing item ${index}:`, itemError.message);
      }
    });

    const result = {
      success: true,
      totalItems: appleProducts.length,
      groupCount: Object.keys(productGroups).length,
      categories: categories,
      productGroups: productGroups,
      debug: {
        totalRows: jsonData.length,
        filteredRows: appleProducts.length,
        headers: headers.slice(0, 10) // Limit for response size
      }
    };

    console.log(`Success: ${result.totalItems} items, ${result.groupCount} groups`);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(result)
    };

  } catch (error) {
    console.error('Error:', error.message);
    
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ 
        error: `Processing failed: ${error.message}`,
        errorType: error.constructor.name
      })
    };
  }
};

// Helper function to get field value from multiple possible column names
function getField(item, possibleNames) {
  for (const name of possibleNames) {
    if (item[name] && String(item[name]).trim() !== '') {
      return String(item[name]).trim();
    }
  }
  return '';
}

// Determine product type
function determineType(model, category) {
  const modelLower = model.toLowerCase();
  const categoryLower = category.toLowerCase();
  
  if (modelLower.includes('macbook pro')) return 'MacBook Pro';
  if (modelLower.includes('macbook air')) return 'MacBook Air';
  if (modelLower.includes('macbook')) return 'MacBook';
  if (modelLower.includes('ipad pro')) return 'iPad Pro';
  if (modelLower.includes('ipad air')) return 'iPad Air';
  if (modelLower.includes('ipad')) return 'iPad';
  if (modelLower.includes('iphone')) return 'iPhone';
  if (modelLower.includes('imac')) return 'iMac';
  if (modelLower.includes('airpod')) return 'AirPods';
  
  // Fallback to category
  if (categoryLower.includes('laptop')) return 'MacBook';
  if (categoryLower.includes('tablet')) return 'iPad';
  if (categoryLower.includes('phone')) return 'iPhone';
  
  return category || 'Apple Product';
}

// Create simple SEO title
function createTitle(productType, processor, storage, memory) {
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