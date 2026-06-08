import { NextResponse } from 'next/server';

export async function POST(request: Request) {
  try {
    const { email, password } = await request.json();

    const expectedEmail = process.env.DASHBOARD_EMAIL;
    const expectedPassword = process.env.DASHBOARD_PASSWORD;

    if (!expectedEmail || !expectedPassword) {
      console.error('ERROR: DASHBOARD_EMAIL and DASHBOARD_PASSWORD must be configured as environment variables.');
      return NextResponse.json({ error: 'Server authentication misconfigured' }, { status: 500 });
    }

    if (email === expectedEmail && password === expectedPassword) {
      // Create a response that sets a mock token in a cookie
      const response = NextResponse.json({ success: true, user: { email } });
      
      // Set simple cookie
      response.cookies.set('dashboard_session', 'authenticated_user_session_token', {
        httpOnly: false, // Set to false so client side script can read it easily for simple routing
        secure: process.env.NODE_ENV === 'production',
        maxAge: 60 * 60 * 24, // 1 day
        path: '/',
      });

      return response;
    }

    return NextResponse.json({ error: 'Invalid email or password' }, { status: 401 });
  } catch (error: any) {
    return NextResponse.json({ error: error.message || 'Internal Server Error' }, { status: 500 });
  }
}
