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
    console.log('🚀 SUPER DEBUG: Function started');
    
    // For Netlify, the file data comes as base64 in event.body
    if (!event.body) {
      console.log('❌ SUPER DEBUG: No event.body found');
      throw new Error('No file data received');
    }

    console.log('🔄 SUPER DEBUG: Processing Excel file...');
    console.log(`📊 SUPER DEBUG: Event body length: ${event.body.length} characters`);
    
    // Decode base64 body to buffer
    let buffer;
    try {
      buffer = Buffer.from(event.body, 'base64');
      console.log(`📊 SUPER DEBUG: Buffer created successfully, size: ${buffer.length} bytes`);
    } catch (bufferError) {
      console.log('❌ SUPER DEBUG: Buffer creation failed:', bufferError);
      throw new Error('Failed to create buffer from base64 data');
    }
    
    // Parse Excel file
    let workbook;
    try {
      console.log('📖 SUPER DEBUG: Attempting to read workbook...');
      workbook = XLSX.read(buffer, { type: 'buffer' });
      console.log('✅ SUPER DEBUG: Workbook read successfully');
    } catch (xlsxError) {
      console.log('❌ SUPER DEBUG: XLSX read failed:', xlsxError);
      throw new Error('Failed to parse Excel file: ' + xlsxError.message);
    }
    
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      console.log('❌ SUPER DEBUG: No sheet names found');
      throw new Error('No sheets found in Excel file');
    }
    
    console.log(`📋 SUPER DEBUG: Found ${workbook.SheetNames.length} sheets: ${workbook.SheetNames.join(', ')}`);
    
    let worksheet, jsonData;
    try {
      worksheet = workbook.Sheets[workbook.SheetNames[0]];
      console.log('✅ SUPER DEBUG: Worksheet accessed successfully');
      
      // Check if worksheet has any data
      const range = worksheet['!ref'];
      console.log(`📏 SUPER DEBUG: Worksheet range: ${range}`);
      
      jsonData = XLSX.utils.sheet_to_json(worksheet);
      console.log(`📈 SUPER DEBUG: JSON conversion successful, rows: ${jsonData.length}`);
    } catch (conversionError) {
      console.log('❌ SUPER DEBUG: Worksheet conversion failed:', conversionError);
      throw new Error('Failed to convert worksheet to JSON: ' + conversionError.message);
    }
    
    // SUPER DEBUG: Even if length is 0, let's see what we got
    console.log('🔍 SUPER DEBUG: Raw jsonData analysis:');
    console.log('Type of jsonData:', typeof jsonData);
    console.log('Is Array:', Array.isArray(jsonData));
    console.log('Length:', jsonData.length);
    
    if (jsonData.length === 0) {
      console.log('⚠️ SUPER DEBUG: Zero rows found, checking worksheet properties...');
      
      // Check worksheet properties
      const worksheetKeys = Object.keys(worksheet);
      console.log('Worksheet keys:', worksheetKeys.slice(0, 20)); // First 20 keys
      
      // Try different parsing options
      console.log('🔄 SUPER DEBUG: Trying alternative parsing methods...');
      
      try {
        // Try with header option
        const jsonWithHeader = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        console.log('With header=1:', jsonWithHeader.length, 'rows');
        if (jsonWithHeader.length > 0) {
          console.log('First row:', jsonWithHeader[0]);
        }
        
        // Try raw parsing
        const jsonRaw = XLSX.utils.sheet_to_json(worksheet, { raw: true });
        console.log('Raw parsing:', jsonRaw.length, 'rows');
        
        // Try different range
        if (worksheet['!ref']) {
          const decoded = XLSX.utils.decode_range(worksheet['!ref']);
          console.log('Range details:', decoded);
        }
        
        // Return debug info instead of error
        return {
          statusCode: 200,
          headers,
          body: JSON.stringify({
            error: 'Zero rows found in Excel file',
            debug: {
              sheetNames: workbook.SheetNames,
              worksheetRange: worksheet['!ref'],
              worksheetKeys: worksheetKeys.slice(0, 50),
              jsonDataLength: jsonData.length,
              alternativeParsing: {
                withHeader: jsonWithHeader.length,
                raw: jsonRaw.length,
                firstRowWithHeader: jsonWithHeader[0] || null
              }
            }
          })
        };
        
      } catch (altError) {
        console.log('❌ Alternative parsing failed:', altError);
      }
    }

    if (!jsonData || jsonData.length === 0) {
      console.log('❌ SUPER DEBUG: Final check - no data found');
      throw new Error('No data found in Excel file after all parsing attempts');
    }

    // If we get here, we have data - continue with original debug logic
    console.log('✅ SUPER DEBUG: Data found, continuing with analysis...');

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

    // Continue with normal processing...
    const allProducts = jsonData.filter(item => {
      const category = (item['Sub-Category'] || '').toString().toLowerCase();
      const model = (item.Model || '').toString().toLowerCase();
      const brand = (item.Brand || '').toString().toLowerCase();
      
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
             model.includes('imac');
    });

    console.log(`✅ Found ${allProducts.length} Apple products after filtering`);

    if (allProducts.length === 0) {
      // Return detailed debug info
      const categories = [...new Set(jsonData.map(item => item['Sub-Category']).filter(Boolean))];
      const brands = [...new Set(jsonData.map(item => item.Brand).filter(Boolean))];
      const models = [...new Set(jsonData.map(item => item.Model).filter(Boolean))].slice(0, 10);
      
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
            sampleData: jsonData.slice(0, 3)
          }
        })
      };
    }

    // Success case - simplified for now
    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        success: true,
        totalItems: allProducts.length,
        message: `Successfully found ${allProducts.length} Apple products!`,
        debug: {
          totalRows: jsonData.length,
          filteredRows: allProducts.length
        }
      })
    };

  } catch (error) {
    console.error('❌ SUPER DEBUG: Final error:', error);
    console.error('❌ SUPER DEBUG: Error stack:', error.stack);
    
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ 
        error: 'Error processing Excel file: ' + error.message,
        stack: error.stack,
        debug: {
          errorType: error.constructor.name,
          errorMessage: error.message
        }
      })
    };
  }
};