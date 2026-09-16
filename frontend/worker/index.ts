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
   TEAMS TOKEN
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
   SUPABASE
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
   MATCHES
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
   EVENT TYPES
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
  start_x?: number | null;
  start_y?: number | null;
  end_x?: number | null;
  end_y?: number | null;
};

type DeleteEventBody = {
  id?: number | string;
};

/* =========================================================
   GET EVENTS
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

  const response =
    await supabaseRequest(
      env,
      `goal_events?select=*&match_id=eq.${encodeURIComponent(
        matchId
      )}&order=minute.asc`
    );

  if (!response.ok) {
    const details =
      await response.text();

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
   CREATE EVENT
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
      null,

    start_x:
      body.start_x ?? null,

    start_y:
      body.start_y ?? null,

    end_x:
      body.end_x ?? null,

    end_y:
      body.end_y ?? null
  };

  const response =
    await supabaseRequest(
      env,
      "goal_events",
      {
        method:
          "POST",

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
        Array.isArray(
          created
        )
          ? created[0]
          : created
    },
    {
      status: 201
    }
  );
}

/* =========================================================
   DELETE EVENT
   ========================================================= */

async function deleteEvent(
  request: Request,
  env: Env
): Promise<Response> {
  let body: DeleteEventBody;

  try {
    body =
      (await request.json()) as DeleteEventBody;
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

  if (
    body.id == null ||
    body.id === ""
  ) {
    return Response.json(
      {
        error:
          "Event-ID fehlt"
      },
      {
        status: 400
      }
    );
  }

  const response =
    await supabaseRequest(
      env,
      `goal_events?id=eq.${encodeURIComponent(
        String(
          body.id
        )
      )}`,
      {
        method:
          "DELETE",

        headers: {
          Prefer:
            "return=representation"
        }
      }
    );

  if (!response.ok) {
    const details =
      await response.text();

    return Response.json(
      {
        error:
          "Event konnte nicht gelöscht werden",
        details
      },
      {
        status:
          response.status
      }
    );
  }

  const deleted =
    await response.json();

  if (
    !Array.isArray(
      deleted
    ) ||
    deleted.length === 0
  ) {
    return Response.json(
      {
        error:
          "Kein Event mit dieser ID gefunden"
      },
      {
        status: 404
      }
    );
  }

  return Response.json({
    success: true,
    deleted:
      deleted[0]
  });
}

/* =========================================================
   EVENTS ROUTE
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

  if (
    request.method ===
    "DELETE"
  ) {
    return deleteEvent(
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
          "GET, POST, DELETE"
      }
    }
  );
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

    if (
      url.pathname ===
      "/api/me"
    ) {
      return handleMe(
        request
      );
    }

    if (
      url.pathname ===
      "/api/matches"
    ) {
      return handleMatches(
        request,
        env
      );
    }

    if (
      url.pathname ===
      "/api/events"
    ) {
      return handleEvents(
        request,
        env
      );
    }

    return env.ASSETS.fetch(
      request
    );
  }
};