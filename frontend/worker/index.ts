import {
  createRemoteJWKSet,
  decodeJwt,
  jwtVerify
} from "jose";

interface Env {
  ASSETS: Fetcher;
}

const TENANT_ID = "1fcb46af-c475-4867-8c22-1ada8dd7cfdf";
const CLIENT_ID = "195954d1-452c-40de-8108-e6baf8a12042";

const APP_ID_URI =
  "api://aka-goals-26-27.mario-demmelbauer.workers.dev/195954d1-452c-40de-8108-e6baf8a12042";

const JWKS_V2 = createRemoteJWKSet(
  new URL(
    `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`
  )
);

const JWKS_V1 = createRemoteJWKSet(
  new URL(
    `https://login.microsoftonline.com/${TENANT_ID}/discovery/keys`
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
    const unverified = decodeJwt(token);

    const version = String(unverified.ver ?? "");
    const audience = String(unverified.aud ?? "");
    const issuer = String(unverified.iss ?? "");
    const tenant = String(unverified.tid ?? "");

    console.log("Teams token claims:", {
      ver: version,
      aud: audience,
      iss: issuer,
      tid: tenant
    });

    let payload;

    if (version === "2.0") {
      const result = await jwtVerify(token, JWKS_V2, {
        issuer: `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
        audience: CLIENT_ID
      });

      payload = result.payload;
    } else {
      const result = await jwtVerify(token, JWKS_V1, {
        issuer: `https://sts.windows.net/${TENANT_ID}/`,
        audience: [CLIENT_ID, APP_ID_URI]
      });

      payload = result.payload;
    }

    if (payload.tid !== TENANT_ID) {
      return Response.json(
        { error: "Invalid tenant" },
        { status: 403 }
      );
    }

    return Response.json({
      authenticated: true,
      debug: {
        ver: payload.ver ?? null,
        aud: payload.aud ?? null,
        iss: payload.iss ?? null,
        tid: payload.tid ?? null
      },
      user: {
        name: payload.name ?? null,
        username:
          payload.preferred_username ??
          payload.upn ??
          payload.unique_name ??
          null,
        objectId: payload.oid ?? null,
        tenantId: payload.tid ?? null
      }
    });
  } catch (error) {
    console.error("Token validation failed:", error);

    let debug = {};

    try {
      const payload = decodeJwt(token);

      debug = {
        ver: payload.ver ?? null,
        aud: payload.aud ?? null,
        iss: payload.iss ?? null,
        tid: payload.tid ?? null
      };
    } catch {
      // Token konnte nicht dekodiert werden
    }

    return Response.json(
      {
        authenticated: false,
        error: "Invalid Teams SSO token",
        debug
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