import {
  useEffect,
  useMemo,
  useState,
  type MouseEvent
} from "react";

import * as microsoftTeams
  from "@microsoft/teams-js";

import "./App.css";


/* =====================================================
   TYPES
   ===================================================== */

type TeamName =
  | "U15"
  | "U16"
  | "U18"
  | "JWR"
  | "Profis";

type EventType =
  | "Tor"
  | "Gegentor";

type ViewMode =
  | "matches"
  | "analysis";

type ApiUser = {
  name?: string | null;
  username?: string | null;
};

type MeResponse = {
  authenticated?: boolean;
  user?: ApiUser;
  error?: string;
};

type MatchItem = {
  id?: number | string;
  team?: string;
  opponent?: string;
  date?: string;
  match_date?: string;
  competition?: string;
  competition_type?: string;

  [key: string]: unknown;
};

type MatchesResponse = {
  success?: boolean;
  matches?: MatchItem[];
  error?: string;
  details?: string;
};

type GoalEvent = {
  id?: number | string;
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

  finish_touch?: string | null;

  set_piece_type?: string | null;

  comment?: string | null;

  [key: string]: unknown;
};

type EventsResponse = {
  success?: boolean;
  events?: GoalEvent[];
  event?: GoalEvent;
  error?: string;
  details?: string;
  deletedId?: number | string;
};

type ScreenPoint = {
  x: number;
  y: number;
};

type DatabasePoint = {
  x: number;
  y: number;
};

type StatItem = {
  name: string;
  count: number;
  percentage: number;
};


/* =====================================================
   CONSTANTS
   ===================================================== */

const TEAMS: TeamName[] = [
  "U15",
  "U16",
  "U18",
  "JWR",
  "Profis"
];

const PHASES = [
  "",
  "Kontrollierter Spielaufbau",
  "Umschalten nach Ballgewinn",
  "Umschalten nach Ballverlust",
  "Standard",
  "Sonstiges"
];

const CREATION_TYPES = [
  "",
  "Steckpass",
  "Cutback",
  "Flanke",
  "Durchbruch",
  "Distanzschuss",
  "Zweiter Ball",
  "Standard",
  "Sonstiges"
];

const FINISH_TOUCHES = [
  "",
  "One Touch",
  "Two Touches",
  ">2 Touches"
];

const SET_PIECE_TYPES = [
  "",
  "Eckball links",
  "Eckball rechts",
  "Elfmeter",
  "Direkter Freistoß",
  "Indirekter Freistoß",
  "Einwurf",
  "Sonstiger Standard"
];


/* =====================================================
   KOORDINATEN
   ===================================================== */

function databaseToScreen(
  databaseX: number,
  databaseY: number
): ScreenPoint {
  return {
    x:
      databaseY,

    y:
      100 -
      databaseX
  };
}


function screenToDatabase(
  screenX: number,
  screenY: number
): DatabasePoint {
  return {
    x:
      100 -
      screenY,

    y:
      screenX
  };
}


function getAssistPoint(
  event: GoalEvent
): DatabasePoint | null {
  if (
    event.assist_x == null ||
    event.assist_y == null
  ) {
    return null;
  }

  return {
    x:
      Number(
        event.assist_x
      ),

    y:
      Number(
        event.assist_y
      )
  };
}


function getFinishPoint(
  event: GoalEvent
): DatabasePoint | null {
  if (
    event.finish_x == null ||
    event.finish_y == null
  ) {
    return null;
  }

  return {
    x:
      Number(
        event.finish_x
      ),

    y:
      Number(
        event.finish_y
      )
  };
}


/* =====================================================
   ZONENLOGIK
   ===================================================== */

/*
 * Feld:
 * Länge 105 m
 * Breite 68 m
 *
 * 16,5-m-Strafraum:
 * 16,5 / 105 = 15,714 %
 *
 * Angriffstor liegt bei x = 100.
 *
 * Beginn Strafraum:
 * 100 - 15,714 = 84,286
 *
 * Strafraumbreite:
 * 40,32 m von 68 m.
 *
 * Linke Grenze:
 * (68 - 40,32) / 2 / 68 * 100
 * ≈ 20,35
 *
 * Rechte Grenze:
 * ≈ 79,65
 */

const FIRST_THIRD =
  100 / 3;

const SECOND_THIRD =
  (100 / 3) * 2;

const PENALTY_BOX_X =
  100 -
  (
    16.5 /
    105
  ) *
    100;

const PENALTY_BOX_Y_MIN =
  (
    (
      68 -
      40.32
    ) /
    2 /
    68
  ) *
  100;

const PENALTY_BOX_Y_MAX =
  100 -
  PENALTY_BOX_Y_MIN;


/*
 * Box wird für die Analyse noch einmal
 * in links / zentral / rechts unterteilt.
 *
 * Die Grenzen 45 / 55 entsprechen
 * den vorhandenen Daten sehr gut:
 *
 * y 38–42  => Box links
 * y 46–50  => Box zentral
 * y 67     => Box rechts
 */

const BOX_LEFT_END =
  45;

const BOX_CENTER_END =
  55;


/*
 * Allgemeine fünf vertikale Korridore:
 *
 * 0–20   linker Flügel
 * 20–40  linker Halbraum
 * 40–60  Zentrum
 * 60–80  rechter Halbraum
 * 80–100 rechter Flügel
 */

function getHorizontalLane(
  y: number
): string {
  if (
    y < 20
  ) {
    return "linker Flügel";
  }

  if (
    y < 40
  ) {
    return "linker Halbraum";
  }

  if (
    y <= 60
  ) {
    return "Zentrum";
  }

  if (
    y <= 80
  ) {
    return "rechter Halbraum";
  }

  return "rechter Flügel";
}


function calculateZone(
  x: number,
  y: number
): string {
  const safeX =
    Math.max(
      0,
      Math.min(
        100,
        x
      )
    );

  const safeY =
    Math.max(
      0,
      Math.min(
        100,
        y
      )
    );


  /* ===================================================
     BOX AM GEGNERISCHEN TOR
     =================================================== */

  const insideAttackingBox =
    safeX >=
      PENALTY_BOX_X &&
    safeY >=
      PENALTY_BOX_Y_MIN &&
    safeY <=
      PENALTY_BOX_Y_MAX;

  if (
    insideAttackingBox
  ) {
    if (
      safeY <
      BOX_LEFT_END
    ) {
      return "Box links";
    }

    if (
      safeY <=
      BOX_CENTER_END
    ) {
      return "Box zentral";
    }

    return "Box rechts";
  }


  /* ===================================================
     ZONE 14
     =================================================== */

  const insideZone14 =
    safeX >=
      SECOND_THIRD &&
    safeX <
      PENALTY_BOX_X &&
    safeY >=
      40 &&
    safeY <=
      60;

  if (
    insideZone14
  ) {
    return "Zone 14 / Zentrum vor Box";
  }


  /* ===================================================
     FELDDRITTEL
     =================================================== */

  const lane =
    getHorizontalLane(
      safeY
    );

  if (
    safeX <
    FIRST_THIRD
  ) {
    return `Aufbaudrittel – ${lane}`;
  }

  if (
    safeX <
    SECOND_THIRD
  ) {
    return `Mitteldrittel – ${lane}`;
  }

  return `Angriffsdrittel – ${lane}`;
}


/* =====================================================
   STATS
   ===================================================== */

function percentage(
  count: number,
  total: number
): number {
  if (
    total <= 0
  ) {
    return 0;
  }

  return Math.round(
    (
      count /
      total
    ) *
      100
  );
}


function createStats(
  events: GoalEvent[],
  getValue:
    (
      event: GoalEvent
    ) => string | null | undefined
): StatItem[] {
  const map =
    new Map<
      string,
      number
    >();

  events.forEach(
    event => {
      const raw =
        getValue(
          event
        );

      const key =
        typeof raw ===
          "string" &&
        raw.trim()
          ? raw.trim()
          : "Nicht angegeben";

      map.set(
        key,
        (
          map.get(
            key
          ) ??
          0
        ) +
          1
      );
    }
  );

  return Array.from(
    map.entries()
  )
    .map(
      (
        [
          name,
          count
        ]
      ) => ({
        name,

        count,

        percentage:
          percentage(
            count,
            events.length
          )
      })
    )
    .sort(
      (
        a,
        b
      ) =>
        b.count -
        a.count
    );
}


/* =====================================================
   EVENT PITCH
   ===================================================== */

function EventPitch({
  events,
  title
}: {
  events: GoalEvent[];
  title: string;
}) {
  return (
    <div className="analysis-pitch-wrapper">

      <h3>
        {title}
      </h3>

      <div className="football-pitch analysis-pitch">

        <div className="pitch-halfway-line" />
        <div className="pitch-center-circle" />
        <div className="penalty-area penalty-area-top" />
        <div className="penalty-area penalty-area-bottom" />
        <div className="goal-area goal-area-top" />
        <div className="goal-area goal-area-bottom" />

        <svg
          className="pitch-line-layer"
          viewBox="0 0 100 100"
          preserveAspectRatio="none"
        >

          {events.map(
            (
              event,
              index
            ) => {
              const assist =
                getAssistPoint(
                  event
                );

              const finish =
                getFinishPoint(
                  event
                );

              if (
                !assist ||
                !finish
              ) {
                return null;
              }

              const a =
                databaseToScreen(
                  assist.x,
                  assist.y
                );

              const f =
                databaseToScreen(
                  finish.x,
                  finish.y
                );

              return (
                <line
                  key={`line-${event.id ?? index}`}
                  x1={a.x}
                  y1={a.y}
                  x2={f.x}
                  y2={f.y}
                  vectorEffect="non-scaling-stroke"
                />
              );
            }
          )}

        </svg>

        {events.map(
          (
            event,
            index
          ) => {
            const assist =
              getAssistPoint(
                event
              );

            if (
              !assist
            ) {
              return null;
            }

            const point =
              databaseToScreen(
                assist.x,
                assist.y
              );

            return (
              <div
                key={`assist-${event.id ?? index}`}
                className="analysis-start-point"
                style={{
                  left:
                    `${point.x}%`,
                  top:
                    `${point.y}%`
                }}
              />
            );
          }
        )}

        {events.map(
          (
            event,
            index
          ) => {
            const finish =
              getFinishPoint(
                event
              );

            if (
              !finish
            ) {
              return null;
            }

            const point =
              databaseToScreen(
                finish.x,
                finish.y
              );

            return (
              <div
                key={`finish-${event.id ?? index}`}
                className="analysis-end-point"
                style={{
                  left:
                    `${point.x}%`,
                  top:
                    `${point.y}%`
                }}
              >
                {event.minute ?? ""}
              </div>
            );
          }
        )}

      </div>

    </div>
  );
}


/* =====================================================
   STAT CARD
   ===================================================== */

function StatCard({
  title,
  stats
}: {
  title: string;
  stats: StatItem[];
}) {
  return (
    <div className="card stat-card">

      <h2>
        {title}
      </h2>

      {stats.length ===
        0 ? (
        <p>
          Keine Daten vorhanden.
        </p>
      ) : (
        <div className="stat-list">

          {stats.map(
            item => (
              <div
                key={
                  item.name
                }
                className="stat-item"
              >

                <div className="stat-header">

                  <span>
                    {item.name}
                  </span>

                  <strong>
                    {item.count}{" "}
                    ({item.percentage}%)
                  </strong>

                </div>

                <div className="stat-bar-track">

                  <div
                    className="stat-bar-fill"
                    style={{
                      width:
                        `${item.percentage}%`
                    }}
                  />

                </div>

              </div>
            )
          )}

        </div>
      )}

    </div>
  );
}


/* =====================================================
   APP
   ===================================================== */

function App() {
  const [
    status,
    setStatus
  ] =
    useState(
      "Teams wird initialisiert …"
    );

  const [
    userName,
    setUserName
  ] =
    useState("");

  const [
    tokenStatus,
    setTokenStatus
  ] =
    useState("");

  const [
    teamsReady,
    setTeamsReady
  ] =
    useState(false);

  const [
    selectedTeam,
    setSelectedTeam
  ] =
    useState<TeamName>(
      "U15"
    );

  const [
    viewMode,
    setViewMode
  ] =
    useState<ViewMode>(
      "matches"
    );

  const [
    matches,
    setMatches
  ] =
    useState<MatchItem[]>(
      []
    );

  const [
    matchesStatus,
    setMatchesStatus
  ] =
    useState("");

  const [
    loadingMatches,
    setLoadingMatches
  ] =
    useState(false);

  const [
    selectedMatch,
    setSelectedMatch
  ] =
    useState<MatchItem | null>(
      null
    );

  const [
    events,
    setEvents
  ] =
    useState<GoalEvent[]>(
      []
    );

  const [
    teamEvents,
    setTeamEvents
  ] =
    useState<GoalEvent[]>(
      []
    );

  const [
    loadingTeamEvents,
    setLoadingTeamEvents
  ] =
    useState(false);

  const [
    showEventForm,
    setShowEventForm
  ] =
    useState(false);

  const [
    eventType,
    setEventType
  ] =
    useState<EventType>(
      "Tor"
    );

  const [
    minute,
    setMinute
  ] =
    useState("");

  const [
    scorer,
    setScorer
  ] =
    useState("");

  const [
    assister,
    setAssister
  ] =
    useState("");

  const [
    phase,
    setPhase
  ] =
    useState("");

  const [
    creationType,
    setCreationType
  ] =
    useState("");

  const [
    finishTouch,
    setFinishTouch
  ] =
    useState("");

  const [
    setPieceType,
    setSetPieceType
  ] =
    useState("");

  const [
    comment,
    setComment
  ] =
    useState("");

  const [
    assistScreenPoint,
    setAssistScreenPoint
  ] =
    useState<ScreenPoint | null>(
      null
    );

  const [
    finishScreenPoint,
    setFinishScreenPoint
  ] =
    useState<ScreenPoint | null>(
      null
    );

  const [
    savingEvent,
    setSavingEvent
  ] =
    useState(false);

  const [
    saveStatus,
    setSaveStatus
  ] =
    useState("");

  const [
    eventToDelete,
    setEventToDelete
  ] =
    useState<GoalEvent | null>(
      null
    );

  const [
    deletingEvent,
    setDeletingEvent
  ] =
    useState(false);

  const [
    deleteStatus,
    setDeleteStatus
  ] =
    useState("");


  /* =====================================================
     TOKEN
     ===================================================== */

  const getTeamsToken =
    async (): Promise<string> => {
      const token =
        await microsoftTeams.authentication.getAuthToken();

      if (
        !token
      ) {
        throw new Error(
          "Kein Teams SSO Token empfangen"
        );
      }

      return token;
    };


  /* =====================================================
     INIT
     ===================================================== */

  useEffect(() => {
    const initTeams =
      async () => {
        try {
          await microsoftTeams.app.initialize();

          const context =
            await microsoftTeams.app.getContext();

          setStatus(
            "Microsoft Teams erkannt"
          );

          const contextUserName =
            context.user?.displayName ??
            context.user?.userPrincipalName ??
            "";

          if (
            contextUserName
          ) {
            setUserName(
              contextUserName
            );
          }

          const token =
            await getTeamsToken();

          const response =
            await fetch(
              "/api/me",
              {
                headers: {
                  Authorization:
                    `Bearer ${token}`
                }
              }
            );

          const data =
            (await response.json()) as MeResponse;

          if (
            !response.ok
          ) {
            setTokenStatus(
              data.error ??
                "SSO Validierung fehlgeschlagen"
            );

            return;
          }

          setUserName(
            data.user?.name ??
            data.user?.username ??
            contextUserName
          );

          setTokenStatus(
            "Teams SSO erfolgreich serverseitig validiert"
          );

          setTeamsReady(
            true
          );
        } catch (
          error
        ) {
          console.error(
            error
          );

          setStatus(
            "AKA Goals läuft außerhalb von Microsoft Teams"
          );
        }
      };

    void initTeams();
  }, []);


  /* =====================================================
     MATCHES
     ===================================================== */

  const loadMatches =
    async (
      team: TeamName
    ) => {
      setLoadingMatches(
        true
      );

      setMatches([]);

      setMatchesStatus(
        `${team}-Spiele werden geladen …`
      );

      try {
        const token =
          await getTeamsToken();

        const response =
          await fetch(
            `/api/matches?team=${encodeURIComponent(
              team
            )}`,
            {
              headers: {
                Authorization:
                  `Bearer ${token}`
              }
            }
          );

        const data =
          (await response.json()) as MatchesResponse;

        if (
          !response.ok
        ) {
          setMatchesStatus(
            data.error ??
              "Spiele konnten nicht geladen werden"
          );

          return;
        }

        const loaded =
          Array.isArray(
            data.matches
          )
            ? data.matches
            : [];

        setMatches(
          loaded
        );

        setMatchesStatus(
          loaded.length === 0
            ? `Keine ${team}-Spiele gefunden`
            : `${loaded.length} ${team}-Spiele geladen`
        );
      } finally {
        setLoadingMatches(
          false
        );
      }
    };


  /* =====================================================
     TEAM EVENTS
     ===================================================== */

  const loadTeamEvents =
    async (
      team: TeamName
    ) => {
      setLoadingTeamEvents(
        true
      );

      try {
        const token =
          await getTeamsToken();

        const response =
          await fetch(
            `/api/team-events?team=${encodeURIComponent(
              team
            )}`,
            {
              headers: {
                Authorization:
                  `Bearer ${token}`
              }
            }
          );

        const data =
          (await response.json()) as EventsResponse;

        if (
          !response.ok
        ) {
          console.error(
            data
          );

          return;
        }

        setTeamEvents(
          Array.isArray(
            data.events
          )
            ? data.events
            : []
        );
      } finally {
        setLoadingTeamEvents(
          false
        );
      }
    };


  useEffect(() => {
    if (
      !teamsReady
    ) {
      return;
    }

    void loadMatches(
      selectedTeam
    );
  }, [teamsReady]);


  /* =====================================================
     MATCH EVENTS
     ===================================================== */

  const loadEvents =
    async (
      match: MatchItem
    ) => {
      if (
        match.id == null
      ) {
        return;
      }

      setSelectedMatch(
        match
      );

      const token =
        await getTeamsToken();

      const response =
        await fetch(
          `/api/events?match_id=${encodeURIComponent(
            String(
              match.id
            )
          )}`,
          {
            headers: {
              Authorization:
                `Bearer ${token}`
            }
          }
        );

      const data =
        (await response.json()) as EventsResponse;

      if (
        response.ok
      ) {
        setEvents(
          Array.isArray(
            data.events
          )
            ? data.events
            : []
        );
      }
    };


  /* =====================================================
     VIEW
     ===================================================== */

  const openMatches =
    () => {
      setViewMode(
        "matches"
      );

      setSelectedMatch(
        null
      );

      setShowEventForm(
        false
      );
    };


  const openAnalysis =
    () => {
      setViewMode(
        "analysis"
      );

      setSelectedMatch(
        null
      );

      setShowEventForm(
        false
      );

      void loadTeamEvents(
        selectedTeam
      );
    };


  /* =====================================================
     FORM
     ===================================================== */

  const openEventForm =
    (
      type: EventType
    ) => {
      setEventType(
        type
      );

      setMinute("");
      setScorer("");
      setAssister("");
      setPhase("");
      setCreationType("");
      setFinishTouch("");
      setSetPieceType("");
      setComment("");

      setAssistScreenPoint(
        null
      );

      setFinishScreenPoint(
        null
      );

      setSaveStatus("");

      setShowEventForm(
        true
      );
    };


  const handlePitchClick =
    (
      event: MouseEvent<HTMLDivElement>
    ) => {
      const rect =
        event.currentTarget.getBoundingClientRect();

      const point:
        ScreenPoint =
        {
          x:
            Math.round(
              (
                (
                  event.clientX -
                  rect.left
                ) /
                rect.width
              ) *
                10000
            ) /
            100,

          y:
            Math.round(
              (
                (
                  event.clientY -
                  rect.top
                ) /
                rect.height
              ) *
                10000
            ) /
            100
        };

      if (
        !assistScreenPoint
      ) {
        setAssistScreenPoint(
          point
        );

        return;
      }

      if (
        !finishScreenPoint
      ) {
        setFinishScreenPoint(
          point
        );

        return;
      }

      setAssistScreenPoint(
        point
      );

      setFinishScreenPoint(
        null
      );
    };


  /* =====================================================
     LIVE ZONEN
     ===================================================== */

  const assistDatabasePoint =
    useMemo(
      () =>
        assistScreenPoint
          ? screenToDatabase(
              assistScreenPoint.x,
              assistScreenPoint.y
            )
          : null,
      [
        assistScreenPoint
      ]
    );


  const finishDatabasePoint =
    useMemo(
      () =>
        finishScreenPoint
          ? screenToDatabase(
              finishScreenPoint.x,
              finishScreenPoint.y
            )
          : null,
      [
        finishScreenPoint
      ]
    );


  const currentAssistZone =
    useMemo(
      () =>
        assistDatabasePoint
          ? calculateZone(
              assistDatabasePoint.x,
              assistDatabasePoint.y
            )
          : "",
      [
        assistDatabasePoint
      ]
    );


  const currentFinishZone =
    useMemo(
      () =>
        finishDatabasePoint
          ? calculateZone(
              finishDatabasePoint.x,
              finishDatabasePoint.y
            )
          : "",
      [
        finishDatabasePoint
      ]
    );


  /* =====================================================
     SAVE
     ===================================================== */

  const saveEvent =
    async () => {
      if (
        !selectedMatch ||
        selectedMatch.id == null ||
        !assistDatabasePoint ||
        !finishDatabasePoint
      ) {
        setSaveStatus(
          "Bitte beide Positionen setzen."
        );

        return;
      }

      setSavingEvent(
        true
      );

      try {
        const token =
          await getTeamsToken();

        const response =
          await fetch(
            "/api/events",
            {
              method:
                "POST",

              headers: {
                Authorization:
                  `Bearer ${token}`,

                "Content-Type":
                  "application/json"
              },

              body:
                JSON.stringify({
                  match_id:
                    selectedMatch.id,

                  team:
                    selectedTeam,

                  event_type:
                    eventType,

                  minute:
                    minute === ""
                      ? null
                      : Number(
                          minute
                        ),

                  scorer:
                    eventType ===
                    "Tor"
                      ? scorer
                      : null,

                  assister:
                    eventType ===
                    "Tor"
                      ? assister
                      : null,

                  phase,

                  creation_type:
                    creationType,

                  finish_touch:
                    finishTouch,

                  set_piece_type:
                    setPieceType,

                  comment,

                  assist_x:
                    assistDatabasePoint.x,

                  assist_y:
                    assistDatabasePoint.y,

                  assist_zone:
                    currentAssistZone,

                  finish_x:
                    finishDatabasePoint.x,

                  finish_y:
                    finishDatabasePoint.y,

                  finish_zone:
                    currentFinishZone
                })
            }
          );

        const data =
          (await response.json()) as EventsResponse;

        if (
          !response.ok
        ) {
          setSaveStatus(
            data.details
              ? `${data.error} – ${data.details}`
              : data.error ??
                  "Speichern fehlgeschlagen"
          );

          return;
        }

        await loadEvents(
          selectedMatch
        );

        setShowEventForm(
          false
        );

        setSaveStatus(
          ""
        );
      } finally {
        setSavingEvent(
          false
        );
      }
    };


  /* =====================================================
     DELETE
     ===================================================== */

  const executeDeleteEvent =
    async () => {
      if (
        eventToDelete?.id == null ||
        !selectedMatch
      ) {
        return;
      }

      setDeletingEvent(
        true
      );

      setDeleteStatus(
        "Ereignis wird gelöscht …"
      );

      try {
        const token =
          await getTeamsToken();

        const response =
          await fetch(
            `/api/events?id=${encodeURIComponent(
              String(
                eventToDelete.id
              )
            )}`,
            {
              method:
                "DELETE",

              headers: {
                Authorization:
                  `Bearer ${token}`
              }
            }
          );

        if (
          response.ok
        ) {
          setEventToDelete(
            null
          );

          setDeleteStatus(
            ""
          );

          await loadEvents(
            selectedMatch
          );
        } else {
          const data =
            (await response.json()) as EventsResponse;

          setDeleteStatus(
            data.error ??
              "Löschen fehlgeschlagen"
          );
        }
      } finally {
        setDeletingEvent(
          false
        );
      }
    };


  /* =====================================================
     TEAM ANALYSIS
     ===================================================== */

  const teamGoals =
    useMemo(
      () =>
        teamEvents.filter(
          event =>
            event.event_type ===
            "Tor"
        ),
      [
        teamEvents
      ]
    );


  const teamConceded =
    useMemo(
      () =>
        teamEvents.filter(
          event =>
            event.event_type ===
            "Gegentor"
        ),
      [
        teamEvents
      ]
    );


  const goalTouchStats =
    useMemo(
      () =>
        createStats(
          teamGoals,
          event =>
            event.finish_touch
        ),
      [
        teamGoals
      ]
    );


  const concededTouchStats =
    useMemo(
      () =>
        createStats(
          teamConceded,
          event =>
            event.finish_touch
        ),
      [
        teamConceded
      ]
    );


  const goalSetPieceStats =
    useMemo(
      () =>
        createStats(
          teamGoals.filter(
            event =>
              Boolean(
                event.set_piece_type
              )
          ),

          event =>
            event.set_piece_type
        ),
      [
        teamGoals
      ]
    );


  const concededSetPieceStats =
    useMemo(
      () =>
        createStats(
          teamConceded.filter(
            event =>
              Boolean(
                event.set_piece_type
              )
          ),

          event =>
            event.set_piece_type
        ),
      [
        teamConceded
      ]
    );


  const goalAssistZoneStats =
    useMemo(
      () =>
        createStats(
          teamGoals.filter(
            event =>
              Boolean(
                event.assist_zone
              )
          ),

          event =>
            event.assist_zone
        ),
      [
        teamGoals
      ]
    );


  const goalFinishZoneStats =
    useMemo(
      () =>
        createStats(
          teamGoals.filter(
            event =>
              Boolean(
                event.finish_zone
              )
          ),

          event =>
            event.finish_zone
        ),
      [
        teamGoals
      ]
    );


  const concededAssistZoneStats =
    useMemo(
      () =>
        createStats(
          teamConceded.filter(
            event =>
              Boolean(
                event.assist_zone
              )
          ),

          event =>
            event.assist_zone
        ),
      [
        teamConceded
      ]
    );


  const concededFinishZoneStats =
    useMemo(
      () =>
        createStats(
          teamConceded.filter(
            event =>
              Boolean(
                event.finish_zone
              )
          ),

          event =>
            event.finish_zone
        ),
      [
        teamConceded
      ]
    );


  /* =====================================================
     HELPERS
     ===================================================== */

  const getMatchTitle =
    (
      match: MatchItem
    ) =>
      typeof match.opponent ===
        "string"
        ? match.opponent
        : "Unbekannter Gegner";


  const goals =
    events.filter(
      event =>
        event.event_type ===
        "Tor"
    );


  const concededGoals =
    events.filter(
      event =>
        event.event_type ===
        "Gegentor"
    );


  const renderEvent =
    (
      event: GoalEvent
    ) => (
      <div
        key={
          String(
            event.id
          )
        }

        className="event-entry"
      >

        <strong>
          {event.minute != null
            ? `${event.minute}. Minute`
            : "Minute unbekannt"}
        </strong>

        {event.scorer && (
          <p>
            Torschütze:{" "}
            {event.scorer}
          </p>
        )}

        {event.assister && (
          <p>
            Assist:{" "}
            {event.assister}
          </p>
        )}

        {event.phase && (
          <p>
            Phase:{" "}
            {event.phase}
          </p>
        )}

        {event.creation_type && (
          <p>
            Entstehung:{" "}
            {event.creation_type}
          </p>
        )}

        {event.finish_touch && (
          <p>
            Abschlusskontakt:{" "}
            {event.finish_touch}
          </p>
        )}

        {event.set_piece_type && (
          <p>
            Standard:{" "}
            {event.set_piece_type}
          </p>
        )}

        {event.assist_zone && (
          <p>
            Assist-Zone:{" "}
            {event.assist_zone}
          </p>
        )}

        {event.finish_zone && (
          <p>
            Abschluss-Zone:{" "}
            {event.finish_zone}
          </p>
        )}

        {event.comment && (
          <p>
            Kommentar:{" "}
            {event.comment}
          </p>
        )}

        <button
          className="delete-event-button"

          onClick={() => {
            setEventToDelete(
              event
            );

            setDeleteStatus(
              ""
            );
          }}
        >
          Löschen
        </button>

      </div>
    );


  /* =====================================================
     UI
     ===================================================== */

  return (
    <div className="app">

      <header>
        <h1>
          AKA Goals
        </h1>

        <span>
          SV Oberbank Ried
        </span>
      </header>

      <main>

        <h2>
          AKA Goals Dashboard
        </h2>

        {userName && (
          <p>
            Angemeldet als{" "}
            <strong>
              {userName}
            </strong>
          </p>
        )}

        <p className="status">
          {status}
        </p>

        <p className="status">
          {tokenStatus}
        </p>


        {/* TEAM AUSWAHL */}

        <div className="cards">

          {TEAMS.map(
            team => (
              <div
                key={
                  team
                }

                className={`card team-card ${
                  selectedTeam ===
                  team
                    ? "active-team"
                    : ""
                }`}

                onClick={() => {
                  setSelectedTeam(
                    team
                  );

                  setSelectedMatch(
                    null
                  );

                  setEvents([]);

                  setTeamEvents([]);

                  setViewMode(
                    "matches"
                  );

                  setShowEventForm(
                    false
                  );

                  void loadMatches(
                    team
                  );
                }}
              >

                <h3>
                  {team}
                </h3>

                <p>
                  Goals &amp; Analysis
                </p>

              </div>
            )
          )}

        </div>


        {/* NAVIGATION */}

        <div className="view-switcher">

          <button
            className={
              viewMode ===
              "matches"
                ? "view-button-active"
                : ""
            }

            onClick={
              openMatches
            }
          >
            Spiele
          </button>

          <button
            className={
              viewMode ===
              "analysis"
                ? "view-button-active"
                : ""
            }

            onClick={
              openAnalysis
            }
          >
            Team-Auswertung
          </button>

        </div>


        {/* TEAM AUSWERTUNG */}

        {viewMode ===
          "analysis" ? (

          <section>

            <h2>
              {selectedTeam} – Team-Auswertung
            </h2>

            {loadingTeamEvents ? (
              <p>
                Auswertung wird geladen …
              </p>
            ) : (
              <>

                <div className="analysis-kpis">

                  <div className="kpi-card">
                    <span>
                      Tore
                    </span>

                    <strong>
                      {teamGoals.length}
                    </strong>
                  </div>

                  <div className="kpi-card">
                    <span>
                      Gegentore
                    </span>

                    <strong>
                      {teamConceded.length}
                    </strong>
                  </div>

                  <div className="kpi-card">
                    <span>
                      Tordifferenz
                    </span>

                    <strong>
                      {teamGoals.length -
                        teamConceded.length}
                    </strong>
                  </div>

                  <div className="kpi-card">
                    <span>
                      Spiele
                    </span>

                    <strong>
                      {matches.length}
                    </strong>
                  </div>

                </div>


                <h2 className="analysis-heading">
                  Abschluss nach Kontakten
                </h2>

                <div className="events-grid">

                  <StatCard
                    title="Tore"
                    stats={
                      goalTouchStats
                    }
                  />

                  <StatCard
                    title="Gegentore"
                    stats={
                      concededTouchStats
                    }
                  />

                </div>


                <h2 className="analysis-heading">
                  Standards
                </h2>

                <div className="events-grid">

                  <StatCard
                    title="Tore nach Standard"
                    stats={
                      goalSetPieceStats
                    }
                  />

                  <StatCard
                    title="Gegentore nach Standard"
                    stats={
                      concededSetPieceStats
                    }
                  />

                </div>


                <h2 className="analysis-heading">
                  Assist-Zonen
                </h2>

                <div className="events-grid">

                  <StatCard
                    title="Tore"
                    stats={
                      goalAssistZoneStats
                    }
                  />

                  <StatCard
                    title="Gegentore"
                    stats={
                      concededAssistZoneStats
                    }
                  />

                </div>


                <h2 className="analysis-heading">
                  Abschluss-Zonen
                </h2>

                <div className="events-grid">

                  <StatCard
                    title="Tore"
                    stats={
                      goalFinishZoneStats
                    }
                  />

                  <StatCard
                    title="Gegentore"
                    stats={
                      concededFinishZoneStats
                    }
                  />

                </div>


                <h2 className="analysis-heading">
                  Positionsdaten
                </h2>

                <div className="analysis-grid">

                  <EventPitch
                    events={
                      teamGoals
                    }

                    title={`Alle Tore (${teamGoals.length})`}
                  />

                  <EventPitch
                    events={
                      teamConceded
                    }

                    title={`Alle Gegentore (${teamConceded.length})`}
                  />

                </div>

              </>
            )}

          </section>

        ) : selectedMatch ? (

          /* EINZELSPIEL */

          <section>

            <button
              onClick={() =>
                setSelectedMatch(
                  null
                )
              }
            >
              ← Zurück zu den Spielen
            </button>

            <h2>
              {selectedTeam}
              {" – "}
              {getMatchTitle(
                selectedMatch
              )}
            </h2>


            <div className="event-actions">

              <button
                onClick={() =>
                  openEventForm(
                    "Tor"
                  )
                }
              >
                + Tor erfassen
              </button>

              <button
                onClick={() =>
                  openEventForm(
                    "Gegentor"
                  )
                }
              >
                + Gegentor erfassen
              </button>

            </div>


            {eventToDelete && (

              <div className="delete-confirm-box">

                <strong>
                  Ereignis löschen?
                </strong>

                <div className="form-actions">

                  <button
                    onClick={() =>
                      void executeDeleteEvent()
                    }

                    disabled={
                      deletingEvent
                    }
                  >
                    Ja, löschen
                  </button>

                  <button
                    onClick={() =>
                      setEventToDelete(
                        null
                      )
                    }
                  >
                    Abbrechen
                  </button>

                </div>

                {deleteStatus && (
                  <p className="status">
                    {deleteStatus}
                  </p>
                )}

              </div>

            )}


            {showEventForm && (

              <div className="card event-form-card">

                <h2>
                  {eventType} erfassen
                </h2>

                <div className="event-form">

                  <label>
                    Minute

                    <input
                      type="number"
                      min="0"
                      max="130"

                      value={
                        minute
                      }

                      onChange={
                        event =>
                          setMinute(
                            event.target.value
                          )
                      }
                    />
                  </label>


                  {eventType ===
                    "Tor" && (
                    <>

                      <label>
                        Torschütze

                        <input
                          value={
                            scorer
                          }

                          onChange={
                            event =>
                              setScorer(
                                event.target.value
                              )
                          }
                        />
                      </label>

                      <label>
                        Assist

                        <input
                          value={
                            assister
                          }

                          onChange={
                            event =>
                              setAssister(
                                event.target.value
                              )
                          }
                        />
                      </label>

                    </>
                  )}


                  <label>
                    Phase

                    <select
                      value={
                        phase
                      }

                      onChange={
                        event =>
                          setPhase(
                            event.target.value
                          )
                      }
                    >
                      {PHASES.map(
                        item => (
                          <option
                            key={
                              item
                            }
                            value={
                              item
                            }
                          >
                            {item ||
                              "Bitte auswählen"}
                          </option>
                        )
                      )}
                    </select>
                  </label>


                  <label>
                    Entstehung

                    <select
                      value={
                        creationType
                      }

                      onChange={
                        event =>
                          setCreationType(
                            event.target.value
                          )
                      }
                    >
                      {CREATION_TYPES.map(
                        item => (
                          <option
                            key={
                              item
                            }
                            value={
                              item
                            }
                          >
                            {item ||
                              "Bitte auswählen"}
                          </option>
                        )
                      )}
                    </select>
                  </label>


                  <label>
                    Abschlusskontakt

                    <select
                      value={
                        finishTouch
                      }

                      onChange={
                        event =>
                          setFinishTouch(
                            event.target.value
                          )
                      }
                    >
                      {FINISH_TOUCHES.map(
                        item => (
                          <option
                            key={
                              item
                            }
                            value={
                              item
                            }
                          >
                            {item ||
                              "Bitte auswählen"}
                          </option>
                        )
                      )}
                    </select>
                  </label>


                  <label>
                    Standard

                    <select
                      value={
                        setPieceType
                      }

                      onChange={
                        event =>
                          setSetPieceType(
                            event.target.value
                          )
                      }
                    >
                      {SET_PIECE_TYPES.map(
                        item => (
                          <option
                            key={
                              item
                            }
                            value={
                              item
                            }
                          >
                            {item ||
                              "Kein Standard"}
                          </option>
                        )
                      )}
                    </select>
                  </label>


                  <label>
                    Kommentar

                    <textarea
                      value={
                        comment
                      }

                      onChange={
                        event =>
                          setComment(
                            event.target.value
                          )
                      }

                      rows={
                        3
                      }
                    />
                  </label>


                  <p className="pitch-help">
                    {!assistScreenPoint
                      ? "1. Klick: Assist / Entstehung"
                      : !finishScreenPoint
                        ? "2. Klick: Abschluss"
                        : "Beide Positionen gesetzt"}
                  </p>


                  <div
                    className="football-pitch"

                    onClick={
                      handlePitchClick
                    }
                  >

                    <div className="pitch-halfway-line" />
                    <div className="pitch-center-circle" />
                    <div className="penalty-area penalty-area-top" />
                    <div className="penalty-area penalty-area-bottom" />
                    <div className="goal-area goal-area-top" />
                    <div className="goal-area goal-area-bottom" />


                    {assistScreenPoint && (
                      <div
                        className="pitch-point assist-point"

                        style={{
                          left:
                            `${assistScreenPoint.x}%`,

                          top:
                            `${assistScreenPoint.y}%`
                        }}
                      >
                        1
                      </div>
                    )}


                    {finishScreenPoint && (
                      <div
                        className="pitch-point finish-point"

                        style={{
                          left:
                            `${finishScreenPoint.x}%`,

                          top:
                            `${finishScreenPoint.y}%`
                        }}
                      >
                        2
                      </div>
                    )}


                    {assistScreenPoint &&
                      finishScreenPoint && (

                      <svg
                        className="pitch-line-layer"
                        viewBox="0 0 100 100"
                        preserveAspectRatio="none"
                      >
                        <line
                          x1={
                            assistScreenPoint.x
                          }
                          y1={
                            assistScreenPoint.y
                          }
                          x2={
                            finishScreenPoint.x
                          }
                          y2={
                            finishScreenPoint.y
                          }
                          vectorEffect="non-scaling-stroke"
                        />
                      </svg>

                    )}

                  </div>


                  {(currentAssistZone ||
                    currentFinishZone) && (

                    <div className="card">

                      <strong>
                        Automatische Zonen
                      </strong>

                      {currentAssistZone && (
                        <p>
                          Assist:
                          {" "}
                          {currentAssistZone}
                        </p>
                      )}

                      {currentFinishZone && (
                        <p>
                          Abschluss:
                          {" "}
                          {currentFinishZone}
                        </p>
                      )}

                    </div>

                  )}


                  <button
                    type="button"

                    onClick={() => {
                      setAssistScreenPoint(
                        null
                      );

                      setFinishScreenPoint(
                        null
                      );
                    }}
                  >
                    Positionen löschen
                  </button>

                </div>


                {saveStatus && (
                  <p className="status">
                    {saveStatus}
                  </p>
                )}


                <div className="form-actions">

                  <button
                    onClick={() =>
                      void saveEvent()
                    }

                    disabled={
                      savingEvent
                    }
                  >
                    {savingEvent
                      ? "Speichern …"
                      : "Speichern"}
                  </button>

                  <button
                    onClick={() =>
                      setShowEventForm(
                        false
                      )
                    }
                  >
                    Abbrechen
                  </button>

                </div>

              </div>

            )}


            <div className="analysis-grid">

              <EventPitch
                events={
                  goals
                }

                title={`Tore (${goals.length})`}
              />

              <EventPitch
                events={
                  concededGoals
                }

                title={`Gegentore (${concededGoals.length})`}
              />

            </div>


            <div className="events-grid">

              <div className="card">

                <h2>
                  Tore ({goals.length})
                </h2>

                {goals.length ===
                  0
                  ? (
                    <p>
                      Keine Tore erfasst.
                    </p>
                  )
                  : goals.map(
                      renderEvent
                    )}

              </div>


              <div className="card">

                <h2>
                  Gegentore ({concededGoals.length})
                </h2>

                {concededGoals.length ===
                  0
                  ? (
                    <p>
                      Keine Gegentore erfasst.
                    </p>
                  )
                  : concededGoals.map(
                      renderEvent
                    )}

              </div>

            </div>

          </section>

        ) : (

          /* SPIELLISTE */

          <section>

            <h2>
              {selectedTeam} Spiele
            </h2>

            <p className="status">
              {matchesStatus}
            </p>

            {loadingMatches
              ? (
                <p>
                  Daten werden geladen …
                </p>
              )
              : matches.map(
                  (
                    match,
                    index
                  ) => (

                    <div
                      key={
                        String(
                          match.id ??
                          index
                        )
                      }

                      className="card match-card"

                      onClick={() =>
                        void loadEvents(
                          match
                        )
                      }
                    >

                      <h3>
                        {getMatchTitle(
                          match
                        )}
                      </h3>

                      <p>
                        Öffnen →
                      </p>

                    </div>

                  )
                )}

          </section>

        )}

      </main>

    </div>
  );
}

export default App;