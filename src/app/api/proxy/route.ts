import { NextResponse } from 'next/server';

export async function POST(request: Request) {
  try {
    const { url, method, headers, body } = await request.json();

    if (!url) {
      return NextResponse.json({ error: 'URL is required' }, { status: 400 });
    }

    const fetchOptions: RequestInit = {
      method: method || 'GET',
      headers: {
        'Content-Type': 'application/json',
        ...headers,
      },
    };

    if (body && ['POST', 'PUT', 'PATCH', 'DELETE'].includes(method)) {
      fetchOptions.body = typeof body === 'string' ? body : JSON.stringify(body);
    }

    const response = await fetch(url, fetchOptions);
    const contentType = response.headers.get('content-type');
    
    let responseData;
    if (contentType && contentType.includes('application/json')) {
      responseData = await response.json();
    } else {
      responseData = { text: await response.text() };
    }

    return NextResponse.json({
      status: response.status,
      data: responseData,
    });
  } catch (error: any) {
    console.error('Proxy error:', error);
    return NextResponse.json({
      status: 500,
      data: { error: error.message || 'Internal Proxy Error' },
    });
  }
}
