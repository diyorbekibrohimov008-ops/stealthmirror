import { handleApiRequest } from '../netlify/lib/router';

async function dispatch(req: Request): Promise<Response> {
  try {
    return await handleApiRequest(req);
  } catch (err) {
    console.error('Unhandled API error in StealthMirror function:', err);
    return Response.json(
      { error: 'Internal server error' },
      { status: 500, headers: { 'cache-control': 'no-store' } }
    );
  }
}

export async function GET(req: Request): Promise<Response> {
  return dispatch(req);
}

export async function POST(req: Request): Promise<Response> {
  return dispatch(req);
}

export async function PUT(req: Request): Promise<Response> {
  return dispatch(req);
}

export async function PATCH(req: Request): Promise<Response> {
  return dispatch(req);
}

export async function DELETE(req: Request): Promise<Response> {
  return dispatch(req);
}
