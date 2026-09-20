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
   TYPES
   ========================================================= */

type AssistPoint = {
  order: number;
  x: number;
  y: number;
  zone?: string | null;
};

type CreateMatchBody = {
  team?: string;
  match_date?: string;
  opponent?: string;
  competition?: string | null;
  home_away?: string | null;
};

type CreateEventBody = {
  match_id?: number | string;

  team?: string;

  event_type?: string;

  minute?: number | null;

  scorer?: string | null;

  assister?: string | null;

  phase?: string | null;

  creation_type?: string | null;

  assist_x?: number | null;
  assist_y?: number | null;

  finish_x?: number | null;
  finish_y?: number | null;

  assist_zone?: string | null;
  finish_zone?: string | null;

  assist_points?: unknown;

  finish_touch?: string | null;

  set_piece_type?: string | null;

  comment?: string | null;
};

type MoveEventBody = {
  team?: string;
  match_id?: number | string;
};


/* =========================================================
   AUTH
   ========================================================= */

async function validateTeamsToken(
  request: Request
): Promise<JWTPayload> {
  const authHeader =
    request.headers.get(
      "Authorization"
    );

  if (
    !authHeader ||
    !authHeader.startsWith(
      "Bearer "
    )
  ) {
    throw new Error(
      "Authorization header missing"
    );
  }

  const token =
    authHeader.substring(
      7
    );

  const unverified =
    decodeJwt(
      token
    );

  const version =
    String(
      unverified.ver ??
      ""
    );


  if (
    version ===
    "2.0"
  ) {
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


async function requireAuth(
  request: Request
): Promise<Response | null> {
  try {
    await validateTeamsToken(
      request
    );

    return null;
  } catch (
    error
  ) {
    console.error(
      "Auth error:",
      error
    );

    return Response.json(
      {
        error:
          "Teams authentication failed"
      },
      {
        status:
          401
      }
    );
  }
}


/* =========================================================
   SUPABASE
   ========================================================= */

function checkSupabaseEnv(
  env: Env
): Response | null {
  if (
    !env.SUPABASE_URL
  ) {
    return Response.json(
      {
        error:
          "SUPABASE_URL fehlt"
      },
      {
        status:
          500
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
        status:
          500
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
   ASSIST POINT VALIDATION
   ========================================================= */

function normalizeAssistPoints(
  value: unknown
): AssistPoint[] {
  if (
    !Array.isArray(
      value
    )
  ) {
    return [];
  }

  const result:
    AssistPoint[] =
    [];

  value.forEach(
    (
      raw,
      index
    ) => {
      if (
        typeof raw !==
          "object" ||
        raw ===
          null
      ) {
        return;
      }

      const item =
        raw as Record<
          string,
          unknown
        >;

      const x =
        Number(
          item.x
        );

      const y =
        Number(
          item.y
        );

      if (
        !Number.isFinite(
          x
        ) ||
        !Number.isFinite(
          y
        )
      ) {
        return;
      }

      if (
        x <
          0 ||
        x >
          100 ||
        y <
          0 ||
        y >
          100
      ) {
        return;
      }

      const zone =
        typeof item.zone ===
          "string"
          ? item.zone.trim()
          : null;

      result.push({
        order:
          index +
          1,

        x,

        y,

        zone:
          zone ||
          null
      });
    }
  );

  return result.slice(
    0,
    3
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
      authenticated:
        true,

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
  } catch (
    error
  ) {
    console.error(
      error
    );

    return Response.json(
      {
        authenticated:
          false,

        error:
          "Invalid Teams SSO token"
      },
      {
        status:
          401
      }
    );
  }
}


/* =========================================================
   MATCHES
   ========================================================= */

async function getMatches(
  request: Request,
  env: Env
): Promise<Response> {
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

  if (
    team
  ) {
    query +=
      `&team=eq.${encodeURIComponent(
        team
      )}`;
  }

  query +=
    "&order=match_date.asc";


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
    success:
      true,

    matches
  });
}


async function createMatch(
  request: Request,
  env: Env,
  user: JWTPayload
): Promise<Response> {
  let body:
    CreateMatchBody;

  try {
    body =
      (await request.json()) as CreateMatchBody;
  } catch {
    return Response.json(
      {
        error:
          "Ungültiger JSON Body"
      },
      {
        status:
          400
      }
    );
  }


  const team =
    body.team?.trim();

  const matchDate =
    body.match_date?.trim();

  const opponent =
    body.opponent?.trim();


  if (
    !team
  ) {
    return Response.json(
      {
        error:
          "Team fehlt"
      },
      {
        status:
          400
      }
    );
  }


  if (
    !matchDate
  ) {
    return Response.json(
      {
        error:
          "Datum fehlt"
      },
      {
        status:
          400
      }
    );
  }


  if (
    !opponent
  ) {
    return Response.json(
      {
        error:
          "Gegner fehlt"
      },
      {
        status:
          400
      }
    );
  }


  const createdBy =
    typeof user.preferred_username ===
      "string"
      ? user.preferred_username
      : typeof user.upn ===
          "string"
        ? user.upn
        : typeof user.name ===
            "string"
          ? user.name
          : null;


  const match = {
    team,

    match_date:
      matchDate,

    opponent,

    competition:
      body.competition?.trim() ||
      null,

    home_away:
      body.home_away?.trim() ||
      null,

    created_by:
      createdBy
  };


  const response =
    await supabaseRequest(
      env,
      "matches",
      {
        method:
          "POST",

        headers: {
          Prefer:
            "return=representation"
        },

        body:
          JSON.stringify(
            match
          )
      }
    );


  if (
    !response.ok
  ) {
    const details =
      await response.text();

    return Response.json(
      {
        error:
          "Spiel konnte nicht angelegt werden",

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
      success:
        true,

      match:
        Array.isArray(
          created
        )
          ? created[0]
          : created
    },
    {
      status:
        201
    }
  );
}


async function deleteMatch(
  request: Request,
  env: Env
): Promise<Response> {
  const url =
    new URL(
      request.url
    );

  const matchId =
    url.searchParams.get(
      "id"
    );


  if (
    !matchId
  ) {
    return Response.json(
      {
        error:
          "Spiel-ID fehlt"
      },
      {
        status:
          400
      }
    );
  }


  const response =
    await supabaseRequest(
      env,

      `matches?id=eq.${encodeURIComponent(
        matchId
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


  if (
    !response.ok
  ) {
    const details =
      await response.text();

    return Response.json(
      {
        error:
          "Spiel konnte nicht gelöscht werden",

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
    deleted.length ===
      0
  ) {
    return Response.json(
      {
        error:
          "Spiel wurde nicht gefunden"
      },
      {
        status:
          404
      }
    );
  }


  return Response.json({
    success:
      true,

    deletedId:
      matchId
  });
}


async function handleMatches(
  request: Request,
  env: Env
): Promise<Response> {
  let user:
    JWTPayload;

  try {
    user =
      await validateTeamsToken(
        request
      );
  } catch (
    error
  ) {
    console.error(
      error
    );

    return Response.json(
      {
        error:
          "Teams authentication failed"
      },
      {
        status:
          401
      }
    );
  }


  const envError =
    checkSupabaseEnv(
      env
    );

  if (
    envError
  ) {
    return envError;
  }


  if (
    request.method ===
    "GET"
  ) {
    return getMatches(
      request,
      env
    );
  }


  if (
    request.method ===
    "POST"
  ) {
    return createMatch(
      request,
      env,
      user
    );
  }


  if (
    request.method ===
    "DELETE"
  ) {
    return deleteMatch(
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
      status:
        405,

      headers: {
        Allow:
          "GET, POST, DELETE"
      }
    }
  );
}


/* =========================================================
   TEAM EVENTS
   ========================================================= */

async function handleTeamEvents(
  request: Request,
  env: Env
): Promise<Response> {
  const authError =
    await requireAuth(
      request
    );

  if (
    authError
  ) {
    return authError;
  }


  const envError =
    checkSupabaseEnv(
      env
    );

  if (
    envError
  ) {
    return envError;
  }


  const url =
    new URL(
      request.url
    );

  const team =
    url.searchParams.get(
      "team"
    );


  if (
    !team
  ) {
    return Response.json(
      {
        error:
          "team fehlt"
      },
      {
        status:
          400
      }
    );
  }


  const response =
    await supabaseRequest(
      env,

      `goal_events?select=*&team=eq.${encodeURIComponent(
        team
      )}&order=minute.asc`
    );


  if (
    !response.ok
  ) {
    const details =
      await response.text();

    return Response.json(
      {
        error:
          "Team-Events konnten nicht geladen werden",

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
    success:
      true,

    events
  });
}


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


  if (
    !matchId
  ) {
    return Response.json(
      {
        error:
          "match_id fehlt"
      },
      {
        status:
          400
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


  if (
    !response.ok
  ) {
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
    success:
      true,

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
  let body:
    CreateEventBody;

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
        status:
          400
      }
    );
  }


  if (
    body.match_id == null ||
    body.match_id ===
      ""
  ) {
    return Response.json(
      {
        error:
          "match_id fehlt"
      },
      {
        status:
          400
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
        status:
          400
      }
    );
  }


  if (
    body.event_type !==
      "Tor" &&
    body.event_type !==
      "Gegentor"
  ) {
    return Response.json(
      {
        error:
          "event_type muss Tor oder Gegentor sein"
      },
      {
        status:
          400
      }
    );
  }


  const assistPoints =
    normalizeAssistPoints(
      body.assist_points
    );


  const lastAssist =
    assistPoints.length >
      0
      ? assistPoints[
          assistPoints.length -
          1
        ]
      : null;


  const event = {
    match_id:
      body.match_id,

    team:
      body.team.trim(),

    event_type:
      body.event_type,

    minute:
      body.minute ??
      null,

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

    /*
     * Legacy-Felder bleiben kompatibel:
     * letzter Assistpunkt vor Abschluss.
     */

    assist_x:
      lastAssist?.x ??
      body.assist_x ??
      null,

    assist_y:
      lastAssist?.y ??
      body.assist_y ??
      null,

    assist_zone:
      lastAssist?.zone ??
      body.assist_zone?.trim() ??
      null,

    /*
     * Alle Voraktionen.
     */

    assist_points:
      assistPoints.length >
        0
        ? assistPoints
        : null,

    finish_x:
      body.finish_x ??
      null,

    finish_y:
      body.finish_y ??
      null,

    finish_zone:
      body.finish_zone?.trim() ||
      null,

    finish_touch:
      body.finish_touch?.trim() ||
      null,

    set_piece_type:
      body.set_piece_type?.trim() ||
      null,

    comment:
      body.comment?.trim() ||
      null
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


  if (
    !response.ok
  ) {
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
      success:
        true,

      event:
        Array.isArray(
          created
        )
          ? created[0]
          : created
    },
    {
      status:
        201
    }
  );
}


/* =========================================================
   MOVE EVENT
   ========================================================= */

async function moveEvent(
  request: Request,
  env: Env
): Promise<Response> {
  const url =
    new URL(
      request.url
    );

  const eventId =
    url.searchParams.get(
      "id"
    );


  if (
    !eventId
  ) {
    return Response.json(
      {
        error:
          "Event-ID fehlt"
      },
      {
        status:
          400
      }
    );
  }


  let body:
    MoveEventBody;

  try {
    body =
      (await request.json()) as MoveEventBody;
  } catch {
    return Response.json(
      {
        error:
          "Ungültiger JSON Body"
      },
      {
        status:
          400
      }
    );
  }


  const targetTeam =
    body.team?.trim();

  const targetMatchId =
    body.match_id;


  if (
    !targetTeam
  ) {
    return Response.json(
      {
        error:
          "Ziel-Team fehlt"
      },
      {
        status:
          400
      }
    );
  }


  if (
    targetMatchId == null ||
    targetMatchId ===
      ""
  ) {
    return Response.json(
      {
        error:
          "Ziel-Spiel fehlt"
      },
      {
        status:
          400
      }
    );
  }


  /*
   * Prüfen, ob das ausgewählte Zielspiel
   * wirklich zum Ziel-Team gehört.
   */

  const targetMatchResponse =
    await supabaseRequest(
      env,

      `matches?select=id,team&id=eq.${encodeURIComponent(
        String(
          targetMatchId
        )
      )}&team=eq.${encodeURIComponent(
        targetTeam
      )}&limit=1`
    );


  if (
    !targetMatchResponse.ok
  ) {
    const details =
      await targetMatchResponse.text();

    return Response.json(
      {
        error:
          "Ziel-Spiel konnte nicht geprüft werden",

        details
      },
      {
        status:
          targetMatchResponse.status
      }
    );
  }


  const targetMatches =
    (await targetMatchResponse.json()) as Array<{
      id?: number | string;
      team?: string;
    }>;


  if (
    !Array.isArray(
      targetMatches
    ) ||
    targetMatches.length ===
      0
  ) {
    return Response.json(
      {
        error:
          "Das Ziel-Spiel gehört nicht zum ausgewählten Team"
      },
      {
        status:
          400
      }
    );
  }


  const response =
    await supabaseRequest(
      env,

      `goal_events?id=eq.${encodeURIComponent(
        eventId
      )}`,

      {
        method:
          "PATCH",

        headers: {
          Prefer:
            "return=representation"
        },

        body:
          JSON.stringify({
            team:
              targetTeam,

            match_id:
              targetMatchId
          })
      }
    );


  if (
    !response.ok
  ) {
    const details =
      await response.text();

    return Response.json(
      {
        error:
          "Ereignis konnte nicht verschoben werden",

        details
      },
      {
        status:
          response.status
      }
    );
  }


  const updated =
    await response.json();


  if (
    !Array.isArray(
      updated
    ) ||
    updated.length ===
      0
  ) {
    return Response.json(
      {
        error:
          "Ereignis wurde nicht gefunden"
      },
      {
        status:
          404
      }
    );
  }


  return Response.json({
    success:
      true,

    event:
      updated[0]
  });
}


/* =========================================================
   DELETE EVENT
   ========================================================= */

async function deleteEvent(
  request: Request,
  env: Env
): Promise<Response> {
  const url =
    new URL(
      request.url
    );

  const eventId =
    url.searchParams.get(
      "id"
    );


  if (
    !eventId
  ) {
    return Response.json(
      {
        error:
          "Event-ID fehlt"
      },
      {
        status:
          400
      }
    );
  }


  const response =
    await supabaseRequest(
      env,

      `goal_events?id=eq.${encodeURIComponent(
        eventId
      )}`,

      {
        method:
          "DELETE",

        headers: {
          Prefer:
            "return=minimal"
        }
      }
    );


  if (
    !response.ok
  ) {
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


  return Response.json({
    success:
      true,

    deletedId:
      eventId
  });
}


/* =========================================================
   EVENTS ROUTE
   ========================================================= */

async function handleEvents(
  request: Request,
  env: Env
): Promise<Response> {
  const authError =
    await requireAuth(
      request
    );

  if (
    authError
  ) {
    return authError;
  }


  const envError =
    checkSupabaseEnv(
      env
    );

  if (
    envError
  ) {
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
    "PATCH"
  ) {
    return moveEvent(
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
      status:
        405,

      headers: {
        Allow:
          "GET, POST, PATCH, DELETE"
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
      "/api/team-events"
    ) {
      return handleTeamEvents(
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
