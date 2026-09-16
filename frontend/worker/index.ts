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

const TENANT_ID =
  "1fcb46af-c475-4867-8c22-1ada8dd7cfdf";

const CLIENT_ID =
  "195954d1-452c-40de-8108-e6baf8a12042";

const APP_ID_URI =
  "api://aka-goals-26-27.mario-demmelbauer.workers.dev/195954d1-452c-40de-8108-e6baf8a12042";

/* =========================================================
   MICROSOFT ENTRA JWKS
   ========================================================= */

const JWKS_V2 =
  createRemoteJWKSet(
    new URL(
      `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`
    )
  );

const JWKS_V1 =
  createRemoteJWKSet(
    new URL(
      `https://login.microsoftonline.com/${TENANT_ID}/discovery/keys`
    )
  );

/* =========================================================
   TEAMS TOKEN VALIDIEREN
   ========================================================= */

async function validateTeamsToken(
  request: Request
): Promise<JWTPayload> {
  const authHeader =
    request.headers.get("Authorization");

  if (
    !authHeader ||
    !authHeader.startsWith("Bearer ")
  ) {
    throw new Error(
      "Authorization header missing"
    );
  }

  const token =
    authHeader.substring(7);

  const unverified =
    decodeJwt(token);

  const version =
    String(
      unverified.ver ?? ""
    );

  /* ---------------------------------------------------------
     TOKEN VERSION 2.0
     --------------------------------------------------------- */

  if (version === "2.0") {
    const result =
      await jwtVerify(
        token,
        JWKS_V2,
        {
          issuer:
            `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,

          audience:
            CLIENT_ID
        }
      );

    if (
      result.payload.tid !==
      TENANT_ID
    ) {
      throw new Error(
        "Invalid tenant"
      );
    }

    return result.payload;
  }

  /* ---------------------------------------------------------
     TOKEN VERSION 1.0
     --------------------------------------------------------- */

  const result =
    await jwtVerify(
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

  if (
    result.payload.tid !==
    TENANT_ID
  ) {
    throw new Error(
      "Invalid tenant"
    );
  }

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
      await validateTeamsToken(
        request
      );

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
        error:
          "Invalid Teams SSO token"
      },
      {
        status: 401
      }
    );
  }
}

/* =========================================================
   SUPABASE HELPER
   ========================================================= */

async function supabaseRequest(
  env: Env,
  path: string,
  init?: RequestInit
): Promise<Response> {
  const headers =
    new Headers(
      init?.headers
    );

  headers.set(
    "apikey",
    env.SUPABASE_SERVICE_ROLE_KEY
  );

  headers.set(
    "Authorization",
    `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`
  );

  headers.set(
    "Content-Type",
    "application/json"
  );

  return fetch(
    `${env.SUPABASE_URL}/rest/v1/${path}`,
    {
      ...init,
      headers
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
  /* ---------------------------------------------------------
     TEAMS AUTHENTIFIZIERUNG
     --------------------------------------------------------- */

  try {
    await validateTeamsToken(
      request
    );
  } catch (error) {
    console.error(
      "Matches auth error:",
      error
    );

    return Response.json(
      {
        error:
          "Teams authentication failed"
      },
      {
        status: 401
      }
    );
  }

  /* ---------------------------------------------------------
     ENV PRÜFEN
     --------------------------------------------------------- */

  if (!env.SUPABASE_URL) {
    return Response.json(
      {
        error:
          "SUPABASE_URL fehlt"
      },
      {
        status: 500
      }
    );
  }

  if (
    !env.SUPABASE_SERVICE_ROLE_KEY
  ) {
    return Response.json(
      {
        error:
          "SUPABASE_SERVICE_ROLE_KEY fehlt"
      },
      {
        status: 500
      }
    );
  }

  /* ---------------------------------------------------------
     SPIELE LADEN
     --------------------------------------------------------- */

  try {
    const url =
      new URL(
        request.url
      );

    const team =
      url.searchParams.get(
        "team"
      );

    let query =
      "matches?select=*";

    if (team) {
      query +=
        `&team=eq.${encodeURIComponent(
          team
        )}`;
    }

    /*
      Optional sortieren.
      Falls die Tabelle eine date-Spalte hat,
      funktioniert das direkt.

      Wenn deine Tabelle anders heißt,
      kann man das später anpassen.
    */

    query +=
      "&order=date.asc";

    const response =
      await supabaseRequest(
        env,
        query
      );

    if (
      !response.ok
    ) {
      const details =
        await response.text();

      console.error(
        "Supabase matches error:",
        details
      );

      return Response.json(
        {
          error:
            "Matches konnten nicht geladen werden",

          details
        },
        {
          status:
            response.status
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
          "Fehler beim Laden der Matches",

        details:
          error instanceof Error
            ? error.message
            : String(error)
      },
      {
        status: 500
      }
    );
  }
}

/* =========================================================
   /api/events
   ========================================================= */

async function handleEvents(
  request: Request,
  env: Env
): Promise<Response> {
  /* ---------------------------------------------------------
     TEAMS AUTHENTIFIZIERUNG
     --------------------------------------------------------- */

  try {
    await validateTeamsToken(
      request
    );
  } catch (error) {
    console.error(
      "Events auth error:",
      error
    );

    return Response.json(
      {
        error:
          "Teams authentication failed"
      },
      {
        status: 401
      }
    );
  }

  /* ---------------------------------------------------------
     ENV PRÜFEN
     --------------------------------------------------------- */

  if (!env.SUPABASE_URL) {
    return Response.json(
      {
        error:
          "SUPABASE_URL fehlt"
      },
      {
        status: 500
      }
    );
  }

  if (
    !env.SUPABASE_SERVICE_ROLE_KEY
  ) {
    return Response.json(
      {
        error:
          "SUPABASE_SERVICE_ROLE_KEY fehlt"
      },
      {
        status: 500
      }
    );
  }

  /* ---------------------------------------------------------
     EVENTS LADEN
     --------------------------------------------------------- */

  try {
    const url =
      new URL(
        request.url
      );

    const matchId =
      url.searchParams.get(
        "match_id"
      );

    if (!matchId) {
      return Response.json(
        {
          error:
            "match_id fehlt"
        },
        {
          status: 400
        }
      );
    }

    const query =
      `goal_events?select=*&match_id=eq.${encodeURIComponent(
        matchId
      )}&order=minute.asc`;

    const response =
      await supabaseRequest(
        env,
        query
      );

    if (
      !response.ok
    ) {
      const details =
        await response.text();

      console.error(
        "Supabase goal_events error:",
        details
      );

      return Response.json(
        {
          error:
            "Events konnten nicht geladen werden",

          details
        },
        {
          status:
            response.status
        }
      );
    }

    const events =
      await response.json();

    return Response.json({
      success: true,
      events
    });
  } catch (error) {
    console.error(
      "Events API error:",
      error
    );

    return Response.json(
      {
        error:
          "Fehler beim Laden der Events",

        details:
          error instanceof Error
            ? error.message
            : String(error)
      },
      {
        status: 500
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
      new URL(
        request.url
      );

    /* -------------------------------------------------------
       API: Benutzer
       ------------------------------------------------------- */

    if (
      url.pathname ===
      "/api/me"
    ) {
      return handleMe(
        request
      );
    }

    /* -------------------------------------------------------
       API: Spiele
       ------------------------------------------------------- */

    if (
      url.pathname ===
      "/api/matches"
    ) {
      return handleMatches(
        request,
        env
      );
    }

    /* -------------------------------------------------------
       API: Tore / Gegentore
       ------------------------------------------------------- */

    if (
      url.pathname ===
      "/api/events"
    ) {
      return handleEvents(
        request,
        env
      );
    }

    /* -------------------------------------------------------
       REACT APP
       ------------------------------------------------------- */

    return env.ASSETS.fetch(
      request
    );
  }
};