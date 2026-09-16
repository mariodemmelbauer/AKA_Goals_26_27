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
   ENTRA JWKS
   ========================================================= */

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
    String(unverified.ver ?? "");

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
   ENV PRÜFEN
   ========================================================= */

function checkSupabaseEnv(
  env: Env
): Response | null {
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

  return null;
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
   /api/matches
   ========================================================= */

async function handleMatches(
  request: Request,
  env: Env
): Promise<Response> {
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

  const envError =
    checkSupabaseEnv(env);

  if (envError) {
    return envError;
  }

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

    const response =
      await supabaseRequest(
        env,
        query
      );

    if (!response.ok) {
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
   EVENT BODY
   ========================================================= */

type CreateEventBody = {
  match_id?: number | string;
  team?: string;
  event_type?: string;

  minute?: number | null;

  scorer?: string | null;
  assister?: string | null;

  phase?: string | null;
  creation_type?: string | null;
};

/* =========================================================
   GET /api/events
   ========================================================= */

async function getEvents(
  request: Request,
  env: Env
): Promise<Response> {
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

  if (!response.ok) {
    const details =
      await response.text();

    console.error(
      "Supabase goal_events GET error:",
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
}

/* =========================================================
   POST /api/events
   ========================================================= */

async function createEvent(
  request: Request,
  env: Env
): Promise<Response> {
  let body: CreateEventBody;

  try {
    body =
      (await request.json()) as CreateEventBody;
  } catch {
    return Response.json(
      {
        error:
          "Ungültiger JSON Body"
      },
      {
        status: 400
      }
    );
  }

  /* ---------------------------------------------------------
     Pflichtfelder
     --------------------------------------------------------- */

  if (
    body.match_id == null ||
    body.match_id === ""
  ) {
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

  if (
    !body.team ||
    !body.team.trim()
  ) {
    return Response.json(
      {
        error:
          "team fehlt"
      },
      {
        status: 400
      }
    );
  }

  if (
    body.event_type !== "Tor" &&
    body.event_type !== "Gegentor"
  ) {
    return Response.json(
      {
        error:
          "event_type muss Tor oder Gegentor sein"
      },
      {
        status: 400
      }
    );
  }

  /* ---------------------------------------------------------
     Minute prüfen
     --------------------------------------------------------- */

  if (
    body.minute != null &&
    (
      !Number.isInteger(
        body.minute
      ) ||
      body.minute < 0 ||
      body.minute > 130
    )
  ) {
    return Response.json(
      {
        error:
          "Minute muss zwischen 0 und 130 liegen"
      },
      {
        status: 400
      }
    );
  }

  /* ---------------------------------------------------------
     Datensatz vorbereiten
     --------------------------------------------------------- */

  const event = {
    match_id:
      body.match_id,

    team:
      body.team.trim(),

    event_type:
      body.event_type,

    minute:
      body.minute ?? null,

    scorer:
      body.scorer?.trim() ||
      null,

    assister:
      body.assister?.trim() ||
      null,

    phase:
      body.phase?.trim() ||
      null,

    creation_type:
      body.creation_type?.trim() ||
      null
  };

  /* ---------------------------------------------------------
     Supabase INSERT
     --------------------------------------------------------- */

  const response =
    await supabaseRequest(
      env,
      "goal_events",
      {
        method: "POST",

        headers: {
          Prefer:
            "return=representation"
        },

        body:
          JSON.stringify(
            event
          )
      }
    );

  if (!response.ok) {
    const details =
      await response.text();

    console.error(
      "Supabase goal_events POST error:",
      details
    );

    return Response.json(
      {
        error:
          "Event konnte nicht gespeichert werden",
        details
      },
      {
        status:
          response.status
      }
    );
  }

  const created =
    await response.json();

  return Response.json(
    {
      success: true,
      event:
        Array.isArray(created)
          ? created[0]
          : created
    },
    {
      status: 201
    }
  );
}

/* =========================================================
   /api/events
   ========================================================= */

async function handleEvents(
  request: Request,
  env: Env
): Promise<Response> {
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

  const envError =
    checkSupabaseEnv(env);

  if (envError) {
    return envError;
  }

  try {
    if (
      request.method ===
      "GET"
    ) {
      return getEvents(
        request,
        env
      );
    }

    if (
      request.method ===
      "POST"
    ) {
      return createEvent(
        request,
        env
      );
    }

    return Response.json(
      {
        error:
          "Method not allowed"
      },
      {
        status: 405,
        headers: {
          Allow:
            "GET, POST"
        }
      }
    );
  } catch (error) {
    console.error(
      "Events API error:",
      error
    );

    return Response.json(
      {
        error:
          "Fehler bei der Event-Verarbeitung",
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

    /* ---------- USER ---------- */

    if (
      url.pathname ===
      "/api/me"
    ) {
      return handleMe(
        request
      );
    }

    /* ---------- MATCHES ---------- */

    if (
      url.pathname ===
      "/api/matches"
    ) {
      return handleMatches(
        request,
        env
      );
    }

    /* ---------- EVENTS ---------- */

    if (
      url.pathname ===
      "/api/events"
    ) {
      return handleEvents(
        request,
        env
      );
    }

    /* ---------- REACT ---------- */

    return env.ASSETS.fetch(
      request
    );
  }
};