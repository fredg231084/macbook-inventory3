const XLSX = require('xlsx');

exports.handler = async (event, context) => {
  // Set timeout to prevent function hanging
  context.callbackWaitsForEmptyEventLoop = false;
  
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers,
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  try {
    console.log('Starting Excel processing...');
    
    if (!event.body) {
      throw new Error('No file data received');
    }

    // Decode and parse Excel
    const buffer = Buffer.from(event.body, 'base64');
    console.log(`Buffer size: ${buffer.length}`);
    
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    console.log(`Sheets: ${workbook.SheetNames.length}`);
    
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Simple parsing - try both methods quickly
    let data = [];
    
    try {
      data = XLSX.utils.sheet_to_json(worksheet);
      console.log(`Method 1: ${data.length} rows`);
    } catch (e) {
      console.log('Method 1 failed, trying method 2');
      const arrayData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      if (arrayData.length > 1) {
        const headers = arrayData[0];
        data = arrayData.slice(1).map(row => {
          const obj = {};
          headers.forEach((header, i) => {
            obj[header] = row[i] || '';
          });
          return obj;
        });
        console.log(`Method 2: ${data.length} rows`);
      }
    }

    if (data.length === 0) {
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          error: 'No data found',
          debug: { sheets: workbook.SheetNames.length }
        })
      };
    }

    console.log(`Processing ${data.length} total rows`);

    // Simple Apple product filter
    const appleProducts = data.filter(item => {
      if (!item || typeof item !== 'object') return false;
      
      const values = Object.values(item).join(' ').toLowerCase();
      return values.includes('apple') || 
             values.includes('macbook') || 
             values.includes('ipad') || 
             values.includes('iphone') ||
             values.includes('imac') ||
             values.includes('laptop') ||
             values.includes('tablet');
    });

    console.log(`Found ${appleProducts.length} Apple products`);

    if (appleProducts.length === 0) {
      // Get sample for debugging
      const sample = data.slice(0, 3).map(item => {
        const keys = Object.keys(item);
        return keys.slice(0, 5).reduce((obj, key) => {
          obj[key] = String(item[key]).substring(0, 50);
          return obj;
        }, {});
      });

      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          error: 'No Apple products found',
          debug: {
            totalRows: data.length,
            sampleData: sample
          }
        })
      };
    }

    // Create simple product groups
    const groups = {};
    const categories = {};

    appleProducts.forEach((item, index) => {
      try {
        // Get product info with fallbacks
        const getField = (names) => {
          for (const name of names) {
            if (item[name]) return String(item[name]).trim();
          }
          return '';
        };

        const model = getField(['Model', 'Product', 'Name']) || `Product ${index + 1}`;
        const category = getField(['Sub-Category', 'Category', 'Type']) || 'Apple Product';
        const processor = getField(['Processor', 'CPU']) || 'Unknown';
        const storage = getField(['Storage', 'SSD']) || 'Unknown';
        const memory = getField(['Memory', 'RAM']) || 'Unknown';

        // Simple product type detection
        let productType = 'Apple Product';
        const modelLower = model.toLowerCase();
        
        if (modelLower.includes('macbook pro')) productType = 'MacBook Pro';
        else if (modelLower.includes('macbook air')) productType = 'MacBook Air';
        else if (modelLower.includes('macbook')) productType = 'MacBook';
        else if (modelLower.includes('ipad pro')) productType = 'iPad Pro';
        else if (modelLower.includes('ipad')) productType = 'iPad';
        else if (modelLower.includes('iphone')) productType = 'iPhone';
        else if (modelLower.includes('imac')) productType = 'iMac';

        // Create group key
        const key = `${productType}_${processor}_${storage}`.replace(/[^a-zA-Z0-9_]/g, '');

        if (!groups[key]) {
          groups[key] = {
            productType,
            processor,
            storage,
            memory,
            seoTitle: `${productType} - ${processor}, ${storage} | Certified Refurbished`,
            basePrice: 999,
            items: [],
            variants: { 'Default-A': { color: 'Default', condition: 'A', quantity: 0 } },
            collections: [productType, 'Apple', 'Refurbished']
          };
        }

        groups[key].items.push({ model, originalData: item });
        groups[key].variants['Default-A'].quantity++;
        categories[productType] = (categories[productType] || 0) + 1;

      } catch (err) {
        console.log(`Error processing item ${index}: ${err.message}`);
      }
    });

    const result = {
      success: true,
      totalItems: appleProducts.length,
      groupCount: Object.keys(groups).length,
      categories,
      productGroups: groups,
      debug: {
        totalRows: data.length,
        filteredRows: appleProducts.length
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
        error: error.message,
        type: error.constructor.name
      })
    };
  }
};
