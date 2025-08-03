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

    console.log('Processing Excel file...');
    
    // Decode base64 body to buffer
    const buffer = Buffer.from(event.body, 'base64');
    
    // Parse Excel file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('No sheets found in Excel file');
    }
    
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    if (!jsonData || jsonData.length === 0) {
      throw new Error('No data found in Excel file');
    }

    console.log(`Found ${jsonData.length} rows in Excel file`);

    // Filter for laptops/MacBooks
    const laptops = jsonData.filter(item => 
      item['Sub-Category'] && 
      item['Sub-Category'].toLowerCase().includes('laptop')
    );

    console.log(`Found ${laptops.length} laptop items`);

    // IMPROVED PRODUCT GROUPING LOGIC
    const productGroups = groupProductsImproved(laptops);

    const processedData = {
      totalItems: laptops.length,
      productGroups: productGroups,
      rawData: laptops,
      groupCount: Object.keys(productGroups).length
    };

    console.log(`Created ${Object.keys(productGroups).length} product groups`);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(processedData)
    };

  } catch (error) {
    console.error('Error processing Excel:', error);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: 'Error processing Excel file: ' + error.message })
    };
  }
};

// IMPROVED GROUPING LOGIC - Creates fewer, more logical product groups
function groupProductsImproved(laptops) {
  const productGroups = {};

  laptops.forEach(item => {
    // Create a more specific product key that groups similar configurations
    const productKey = createImprovedProductKey(item);

    if (!productGroups[productKey]) {
      productGroups[productKey] = {
        model: normalizeModel(item.Model),
        processor: normalizeProcessor(item.Processor),
        storage: normalizeStorage(item.Storage),
        memory: normalizeMemory(item.Memory),
        basePrice: extractPrice(item),
        items: [],
        variants: {}
      };
    }

    // Group by condition and color for variants
    const variantKey = `${item.Color || 'Default'}_${item.Condition || 'Unknown'}`;
    
    if (!productGroups[productKey].variants[variantKey]) {
      productGroups[productKey].variants[variantKey] = {
        color: item.Color || 'Default',
        condition: item.Condition || 'Unknown',
        quantity: 0,
        items: []
      };
    }

    productGroups[productKey].variants[variantKey].quantity++;
    productGroups[productKey].variants[variantKey].items.push(item);
    productGroups[productKey].items.push(item);
  });

  return productGroups;
}

function createImprovedProductKey(item) {
  const model = normalizeModel(item.Model || 'Unknown');
  const processor = normalizeProcessor(item.Processor || '');
  const storage = normalizeStorage(item.Storage || '');
  const memory = normalizeMemory(item.Memory || '');

  // Create a more logical grouping key
  return `${model}_${processor}_${storage}_${memory}`;
}

function normalizeModel(model) {
  if (!model) return 'Unknown';
  
  // Extract main model info
  const modelStr = model.toString().toLowerCase();
  
  if (modelStr.includes('macbook pro')) return 'MacBook Pro';
  if (modelStr.includes('macbook air')) return 'MacBook Air';
  if (modelStr.includes('macbook')) return 'MacBook';
  if (modelStr.includes('imac')) return 'iMac';
  if (modelStr.includes('mac mini')) return 'Mac Mini';
  if (modelStr.includes('mac studio')) return 'Mac Studio';
  
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
  if (procStr.includes('intel') || procStr.includes('i7')) return 'Intel i7';
  if (procStr.includes('i5')) return 'Intel i5';
  if (procStr.includes('i3')) return 'Intel i3';
  
  return processor.substring(0, 15);
}

function normalizeStorage(storage) {
  if (!storage) return 'Unknown';
  
  const storageStr = storage.toString().toLowerCase();
  
  if (storageStr.includes('8tb')) return '8TB';
  if (storageStr.includes('4tb')) return '4TB';
  if (storageStr.includes('2tb')) return '2TB';
  if (storageStr.includes('1tb') || storageStr.includes('1000gb')) return '1TB';
  if (storageStr.includes('512gb')) return '512GB';
  if (storageStr.includes('256gb')) return '256GB';
  if (storageStr.includes('128gb')) return '128GB';
  if (storageStr.includes('64gb')) return '64GB';
  
  return storage;
}

function normalizeMemory(memory) {
  if (!memory) return 'Unknown';
  
  const memStr = memory.toString().toLowerCase();
  
  if (memStr.includes('128gb')) return '128GB';
  if (memStr.includes('64gb')) return '64GB';
  if (memStr.includes('32gb')) return '32GB';
  if (memStr.includes('16gb')) return '16GB';
  if (memStr.includes('8gb')) return '8GB';
  if (memStr.includes('4gb')) return '4GB';
  
  return memory;
}

function extractPrice(item) {
  // Try to extract price from various possible fields
  const priceFields = ['Price', 'Cost', 'Value', 'Amount', 'price', 'cost', 'value', 'amount'];
  
  for (const field of priceFields) {
    if (item[field] && !isNaN(parseFloat(item[field]))) {
      const price = parseFloat(item[field]);
      // Basic validation - price should be reasonable for a laptop
      if (price > 50 && price < 50000) {
        return price;
      }
    }
  }
  
  // Default price based on specs if no price found
  return estimatePrice(item);
}

function estimatePrice(item) {
  let basePrice = 500; // Starting base price
  
  const model = (item.Model || '').toLowerCase();
  const processor = (item.Processor || '').toLowerCase();
  const storage = (item.Storage || '').toLowerCase();
  const memory = (item.Memory || '').toLowerCase();
  
  // Model adjustments
  if (model.includes('macbook pro')) basePrice += 500;
  if (model.includes('macbook air')) basePrice += 200;
  
  // Processor adjustments
  if (processor.includes('m3')) basePrice += 800;
  if (processor.includes('m2')) basePrice += 600;
  if (processor.includes('m1')) basePrice += 400;
  if (processor.includes('pro') || processor.includes('max')) basePrice += 500;
  
  // Storage adjustments
  if (storage.includes('1tb')) basePrice += 300;
  if (storage.includes('512gb')) basePrice += 150;
  if (storage.includes('2tb')) basePrice += 600;
  
  // Memory adjustments
  if (memory.includes('32gb')) basePrice += 400;
  if (memory.includes('16gb')) basePrice += 200;
  if (memory.includes('64gb')) basePrice += 800;
  
  return basePrice;
}