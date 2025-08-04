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
    console.log('🧪 MINIMAL TEST: Function called successfully');
    console.log('📊 Body length:', event.body ? event.body.length : 0);
    console.log('📋 Headers:', JSON.stringify(event.headers));
    
    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        success: true,
        message: '🎉 Function is working! Ready for Excel processing.',
        debug: {
          bodyReceived: !!event.body,
          bodyLength: event.body ? event.body.length : 0,
          httpMethod: event.httpMethod,
          timestamp: new Date().toISOString()
        }
      })
    };

  } catch (error) {
    console.error('❌ MINIMAL TEST ERROR:', error);
    
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ 
        error: 'Test failed: ' + error.message,
        errorType: error.constructor.name
      })
    };
  }
};