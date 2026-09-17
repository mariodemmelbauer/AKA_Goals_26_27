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

            const p =
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
                    `${p.x}%`,
                  top:
                    `${p.y}%`
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

            const p =
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
                    `${p.x}%`,
                  top:
                    `${p.y}%`
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
     EVENTS EINZELSPIEL
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

  /* =====================================================
     EVENT FORM
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

  const saveEvent =
    async () => {
      if (
        !selectedMatch ||
        selectedMatch.id == null ||
        !assistScreenPoint ||
        !finishScreenPoint
      ) {
        setSaveStatus(
          "Bitte beide Positionen setzen."
        );

        return;
      }

      const assistDatabase =
        screenToDatabase(
          assistScreenPoint.x,
          assistScreenPoint.y
        );

      const finishDatabase =
        screenToDatabase(
          finishScreenPoint.x,
          finishScreenPoint.y
        );

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
                    eventType === "Tor"
                      ? scorer
                      : null,

                  assister:
                    eventType === "Tor"
                      ? assister
                      : null,

                  phase,

                  creation_type:
                    creationType,

                  assist_x:
                    assistDatabase.x,

                  assist_y:
                    assistDatabase.y,

                  finish_x:
                    finishDatabase.x,

                  finish_y:
                    finishDatabase.y
                })
            }
          );

        const data =
          (await response.json()) as EventsResponse;

        if (
          !response.ok
        ) {
          setSaveStatus(
            data.error ??
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

          await loadEvents(
            selectedMatch
          );
        }
      } finally {
        setDeletingEvent(
          false
        );
      }
    };

  /* =====================================================
     ANALYSIS DATA
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

  const phaseStats =
    useMemo(
      () => {
        const map =
          new Map<
            string,
            number
          >();

        teamEvents.forEach(
          event => {
            const key =
              event.phase ||
              "Nicht angegeben";

            map.set(
              key,
              (
                map.get(
                  key
                ) ||
                0
              ) +
                1
            );
          }
        );

        return Array.from(
          map.entries()
        ).sort(
          (
            a,
            b
          ) =>
            b[1] -
            a[1]
        );
      },
      [
        teamEvents
      ]
    );

  const creationStats =
    useMemo(
      () => {
        const map =
          new Map<
            string,
            number
          >();

        teamEvents.forEach(
          event => {
            const key =
              event.creation_type ||
              "Nicht angegeben";

            map.set(
              key,
              (
                map.get(
                  key
                ) ||
                0
              ) +
                1
            );
          }
        );

        return Array.from(
          map.entries()
        ).sort(
          (
            a,
            b
          ) =>
            b[1] -
            a[1]
        );
      },
      [
        teamEvents
      ]
    );

  const minuteBuckets =
    useMemo(
      () => {
        const buckets =
          selectedTeam === "U15"
            ? [
                {
                  label:
                    "0–15",
                  min:
                    0,
                  max:
                    15
                },
                {
                  label:
                    "16–30",
                  min:
                    16,
                  max:
                    30
                },
                {
                  label:
                    "31–40",
                  min:
                    31,
                  max:
                    40
                },
                {
                  label:
                    "41–55",
                  min:
                    41,
                  max:
                    55
                },
                {
                  label:
                    "56–70",
                  min:
                    56,
                  max:
                    70
                },
                {
                  label:
                    "71–85",
                  min:
                    71,
                  max:
                    85
                }
              ]
            : [
                {
                  label:
                    "0–15",
                  min:
                    0,
                  max:
                    15
                },
                {
                  label:
                    "16–30",
                  min:
                    16,
                  max:
                    30
                },
                {
                  label:
                    "31–45",
                  min:
                    31,
                  max:
                    45
                },
                {
                  label:
                    "46–60",
                  min:
                    46,
                  max:
                    60
                },
                {
                  label:
                    "61–75",
                  min:
                    61,
                  max:
                    75
                },
                {
                  label:
                    "76–95",
                  min:
                    76,
                  max:
                    95
                }
              ];

        return buckets.map(
          bucket => {
            const goals =
              teamGoals.filter(
                event =>
                  event.minute != null &&
                  event.minute >=
                    bucket.min &&
                  event.minute <=
                    bucket.max
              ).length;

            const conceded =
              teamConceded.filter(
                event =>
                  event.minute != null &&
                  event.minute >=
                    bucket.min &&
                  event.minute <=
                    bucket.max
              ).length;

            return {
              ...bucket,
              goals,
              conceded
            };
          }
        );
      },
      [
        selectedTeam,
        teamGoals,
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

                  setViewMode(
                    "matches"
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

        <div className="event-actions">

          <button
            onClick={
              openMatches
            }
          >
            Spiele
          </button>

          <button
            onClick={
              openAnalysis
            }
          >
            Team-Auswertung
          </button>

        </div>

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

                <div className="cards">

                  <div className="card">
                    <h3>
                      Tore
                    </h3>

                    <strong>
                      {teamGoals.length}
                    </strong>
                  </div>

                  <div className="card">
                    <h3>
                      Gegentore
                    </h3>

                    <strong>
                      {teamConceded.length}
                    </strong>
                  </div>

                  <div className="card">
                    <h3>
                      Ereignisse
                    </h3>

                    <strong>
                      {teamEvents.length}
                    </strong>
                  </div>

                </div>

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

                <div className="events-grid">

                  <div className="card">

                    <h2>
                      Spielphasen
                    </h2>

                    {phaseStats.map(
                      (
                        [
                          name,
                          count
                        ]
                      ) => (
                        <p
                          key={
                            name
                          }
                        >
                          <strong>
                            {name}:
                          </strong>{" "}
                          {count}
                        </p>
                      )
                    )}

                  </div>

                  <div className="card">

                    <h2>
                      Entstehung
                    </h2>

                    {creationStats.map(
                      (
                        [
                          name,
                          count
                        ]
                      ) => (
                        <p
                          key={
                            name
                          }
                        >
                          <strong>
                            {name}:
                          </strong>{" "}
                          {count}
                        </p>
                      )
                    )}

                  </div>

                </div>

                <div className="card">

                  <h2>
                    Tore nach Spielminute
                  </h2>

                  <div className="minute-table">

                    {minuteBuckets.map(
                      bucket => (
                        <div
                          key={
                            bucket.label
                          }
                          className="minute-row"
                        >
                          <strong>
                            {bucket.label}
                          </strong>

                          <span>
                            Tore:{" "}
                            {bucket.goals}
                          </span>

                          <span>
                            Gegentore:{" "}
                            {bucket.conceded}
                          </span>
                        </div>
                      )
                    )}

                  </div>

                </div>

              </>
            )}

          </section>

        ) : selectedMatch ? (

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

              </div>

            )}

            {showEventForm && (

              <div className="card event-form-card">

                <h3>
                  {eventType} erfassen
                </h3>

                <div className="event-form">

                  <input
                    type="number"
                    placeholder="Minute"
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

                  {eventType ===
                    "Tor" && (
                    <>
                      <input
                        placeholder="Torschütze"
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

                      <input
                        placeholder="Assist"
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
                    </>
                  )}

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
                            "Phase"}
                        </option>
                      )
                    )}
                  </select>

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
                            "Entstehung"}
                        </option>
                      )
                    )}
                  </select>

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

                  </div>

                </div>

                {saveStatus && (
                  <p>
                    {saveStatus}
                  </p>
                )}

                <button
                  onClick={() =>
                    void saveEvent()
                  }
                  disabled={
                    savingEvent
                  }
                >
                  Speichern
                </button>

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

          </section>

        ) : (

          <section>

            <h2>
              {selectedTeam} Spiele
            </h2>

            <p>
              {matchesStatus}
            </p>

            {!loadingMatches &&
              matches.map(
                (
                  match,
                  index
                ) => (
                  <div
                    key={String(
                      match.id ??
                        index
                    )}
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