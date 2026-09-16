import {
  createRemoteJWKSet,
  decodeJwt,
  jwtVerify,
  type JWTPayload
} from "jose";

interface Env {
  ASSETS: Fetcher;

  SUPABASE_URL: string;
  SUPABASE_SERVICE_ROLE_KEY: string;
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


/* =========================================================
   TEAMS TOKEN PRÜFEN
   ========================================================= */

async function validateTeamsToken(
  request: Request
): Promise<JWTPayload> {

  const authHeader = request.headers.get("Authorization");

  if (!authHeader?.startsWith("Bearer ")) {
    throw new Error("Authorization header missing");
  }

  const token = authHeader.substring(7);

  const unverified = decodeJwt(token);

  const version = String(unverified.ver ?? "");

  if (version === "2.0") {

    const result = await jwtVerify(
      token,
      JWKS_V2,
      {
        issuer:
          `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,

        audience: CLIENT_ID
      }
    );

    return result.payload;
  }

  const result = await jwtVerify(
    token,
    JWKS_V1,
    {
      issuer:
        `https://sts.windows.net/${TENANT_ID}/`,

      audience: [
        CLIENT_ID,
        APP_ID_URI
      ]
    }
  );

  return result.payload;
}


/* =========================================================
   /api/me
   ========================================================= */

async function handleMe(
  request: Request
): Promise<Response> {

  try {

    const payload =
      await validateTeamsToken(request);

    if (payload.tid !== TENANT_ID) {

      return Response.json(
        {
          authenticated: false,
          error: "Invalid tenant"
        },
        {
          status: 403
        }
      );
    }

    return Response.json({
      authenticated: true,

      user: {
        name:
          payload.name ??
          null,

        username:
          payload.preferred_username ??
          payload.upn ??
          payload.unique_name ??
          null,

        objectId:
          payload.oid ??
          null,

        tenantId:
          payload.tid ??
          null
      }
    });

  } catch (error) {

    console.error(
      "Teams token validation failed:",
      error
    );

    return Response.json(
      {
        authenticated: false,
        error: "Invalid Teams SSO token"
      },
      {
        status: 401
      }
    );
  }
}


/* =========================================================
   SUPABASE
   ========================================================= */

async function supabaseRequest(
  env: Env,
  path: string
): Promise<Response> {

  return fetch(
    `${env.SUPABASE_URL}/rest/v1/${path}`,
    {
      headers: {
        apikey:
          env.SUPABASE_SERVICE_ROLE_KEY,

        Authorization:
          `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`,

        "Content-Type":
          "application/json"
      }
    }
  );
}


/* =========================================================
   /api/matches
   ========================================================= */

async function handleMatches(
  request: Request,
  env: Env
): Promise<Response> {

  try {

    /*
      Erst Teams SSO validieren.
      Ohne gültige Teams-Anmeldung gibt es keine Daten.
    */

    await validateTeamsToken(request);

    const url =
      new URL(request.url);

    const team =
      url.searchParams.get("team");

    let query =
      "matches?select=*";

    if (team) {

      query +=
        `&team=eq.${encodeURIComponent(team)}`;
    }

    /*
      Vorerst holen wir alle Spalten.
      Damit sehen wir exakt die bestehende
      matches-Struktur aus dem Streamlit-Projekt.
    */

    const response =
      await supabaseRequest(
        env,
        query
      );

    if (!response.ok) {

      const errorText =
        await response.text();

      console.error(
        "Supabase matches error:",
        errorText
      );

      return Response.json(
        {
          error:
            "Matches konnten nicht geladen werden",

          details:
            errorText
        },
        {
          status: 500
        }
      );
    }

    const matches =
      await response.json();

    return Response.json({
      success: true,
      matches
    });

  } catch (error) {

    console.error(
      "Matches API error:",
      error
    );

    return Response.json(
      {
        error:
          "Unauthorized"
      },
      {
        status: 401
      }
    );
  }
}


/* =========================================================
   WORKER
   ========================================================= */

export default {

  async fetch(
    request: Request,
    env: Env
  ): Promise<Response> {

    const url =
      new URL(request.url);


    /* ---------- Benutzer ---------- */

    if (
      url.pathname === "/api/me"
    ) {

      return handleMe(request);
    }


    /* ---------- Spiele ---------- */

    if (
      url.pathname === "/api/matches"
    ) {

      return handleMatches(
        request,
        env
      );
    }


    /* ---------- React ---------- */

    return env.ASSETS.fetch(
      request
    );
  }
};