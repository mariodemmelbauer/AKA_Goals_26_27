import { createRemoteJWKSet, jwtVerify } from "jose";

interface Env {
  ASSETS: Fetcher;
}

const TENANT_ID = "1fcb46af-c475-4867-8c22-1ada8dd7cfdf";

const CLIENT_ID = "195954d1-452c-40de-8108-e6baf8a12042";

const APP_ID_URI =
  "api://aka-goals-26-27.mario-demmelbauer.workers.dev/195954d1-452c-40de-8108-e6baf8a12042";

const JWKS = createRemoteJWKSet(
  new URL(
    `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`
  )
);

async function handleMe(request: Request): Promise<Response> {
  const authHeader = request.headers.get("Authorization");

  if (!authHeader?.startsWith("Bearer ")) {
    return Response.json(
      { error: "Authorization header missing" },
      { status: 401 }
    );
  }

  const token = authHeader.substring(7);

  try {
    const { payload } = await jwtVerify(token, JWKS, {
      issuer: `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
      audience: APP_ID_URI
    });

    if (payload.tid !== TENANT_ID) {
      return Response.json(
        { error: "Invalid tenant" },
        { status: 403 }
      );
    }

    return Response.json({
      authenticated: true,

      user: {
        name: payload.name ?? null,
        username:
          payload.preferred_username ??
          payload.upn ??
          null,

        objectId: payload.oid ?? null,
        tenantId: payload.tid ?? null
      }
    });
  } catch (error) {
    console.error("Token validation failed:", error);

    return Response.json(
      {
        authenticated: false,
        error: "Invalid Teams SSO token"
      },
      { status: 401 }
    );
  }
}

export default {
  async fetch(
    request: Request,
    env: Env
  ): Promise<Response> {

    const url = new URL(request.url);

    if (url.pathname === "/api/me") {
      return handleMe(request);
    }

    return env.ASSETS.fetch(request);
  }
};