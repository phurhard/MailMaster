// Remember to run `npm install` to install the new dependencies.

import express, { Request, Response } from 'express';
import cookieParser from 'cookie-parser';
import dotenv from 'dotenv';
import { URLSearchParams } from 'url';

dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

app.use(cookieParser());

const SUPABASE_URL = 'https://api.supabase.com';

app.get('/', (req: Request, res: Response) => {
  res.send('<a href="/connect-supabase">Connect Supabase</a>');
});

app.get('/connect-supabase', (req: Request, res: Response) => {
  const params = new URLSearchParams({
    response_type: 'code',
    client_id: process.env.SUPABASE_OAUTH_CLIENT_ID!,
    redirect_uri: process.env.REDIRECT_URI!
    // You can also add a state parameter for CSRF protection
  });

  const authorizationUrl = `${SUPABASE_URL}/v1/oauth/authorize?${params.toString()}`;
  res.redirect(authorizationUrl);
});

app.get('/auth/callback', async (req: Request, res: Response) => {
  const { code } = req.query;

  if (typeof code !== 'string') {
    return res.status(400).send('Invalid authorization code.');
  }

  try {
    const tokenUrl = `${SUPABASE_URL}/v1/oauth/token`;
    const tokenParams = new URLSearchParams({
      grant_type: 'authorization_code',
      client_id: process.env.SUPABASE_OAUTH_CLIENT_ID!,
      client_secret: process.env.SUPABASE_OAUTH_CLIENT_SECRET!,
      redirect_uri: process.env.REDIRECT_URI!,
      code,
    });

    const tokenResponse = await fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: tokenParams,
    });

    if (!tokenResponse.ok) {
      const errorText = await tokenResponse.text();
      throw new Error(`Failed to exchange authorization code for token: ${errorText}`);
    }

    const { access_token, refresh_token } = await tokenResponse.json();

    // TODO: Store the access_token and refresh_token in your database
    // associated with the user.
    console.log('Access Token:', access_token);
    console.log('Refresh Token:', refresh_token);

    res.redirect('/success');
  } catch (error) {
    console.error(error);
    res.redirect('/error');
  }
});

app.get('/success', (req: Request, res: Response) => {
  res.send('<h1>Supabase integration connected successfully!</h1>');
});

app.get('/error', (req: Request, res: Response) => {
  res.send('<h1>Failed to connect Supabase integration.</h1>');
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});