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

type StatItem = {
  name: string;
  count: number;
  percentage: number;
};

type HeatCell = {
  row: number;
  col: number;
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


/* =====================================================
   KOORDINATEN

   Datenbank:
   x = Spielfeldlänge
   0   = eigenes Tor
   100 = gegnerisches Tor

   y = Spielfeldbreite
   0   = links
   100 = rechts

   Bildschirm:
   horizontal = DB-y
   vertikal   = 100 - DB-x
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


/* =====================================================
   EVENT POINTS
   ===================================================== */

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
   PROZENT
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


/* =====================================================
   STATISTIK
   ===================================================== */

function createStats(
  events: GoalEvent[],
  field:
    | "phase"
    | "creation_type"
): StatItem[] {
  const map =
    new Map<
      string,
      number
    >();

  events.forEach(
    event => {
      const raw =
        event[field];

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
   HEATMAP
   ===================================================== */

function buildHeatmap(
  events: GoalEvent[],
  rows = 12,
  cols = 8
): HeatCell[] {
  const counts =
    Array.from(
      {
        length:
          rows * cols
      },
      () => 0
    );

  let positionedEvents =
    0;

  events.forEach(
    event => {
      const finish =
        getFinishPoint(
          event
        );

      if (
        !finish
      ) {
        return;
      }

      const screen =
        databaseToScreen(
          finish.x,
          finish.y
        );

      const safeX =
        Math.min(
          99.999,
          Math.max(
            0,
            screen.x
          )
        );

      const safeY =
        Math.min(
          99.999,
          Math.max(
            0,
            screen.y
          )
        );

      const col =
        Math.floor(
          (
            safeX /
            100
          ) *
            cols
        );

      const row =
        Math.floor(
          (
            safeY /
            100
          ) *
            rows
        );

      const index =
        row *
          cols +
        col;

      counts[index] +=
        1;

      positionedEvents +=
        1;
    }
  );

  return counts.map(
    (
      count,
      index
    ) => ({
      row:
        Math.floor(
          index /
            cols
        ),

      col:
        index %
        cols,

      count,

      percentage:
        percentage(
          count,
          positionedEvents
        )
    })
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
                {event.minute ??
                  ""}
              </div>
            );
          }
        )}

      </div>

    </div>
  );
}


/* =====================================================
   FINISH HEATMAP
   ===================================================== */

function FinishHeatmap({
  events,
  title
}: {
  events: GoalEvent[];
  title: string;
}) {
  const rows =
    12;

  const cols =
    8;

  const cells =
    useMemo(
      () =>
        buildHeatmap(
          events,
          rows,
          cols
        ),
      [
        events
      ]
    );

  const maxCount =
    Math.max(
      0,
      ...cells.map(
        cell =>
          cell.count
      )
    );

  const positionedCount =
    events.filter(
      event =>
        getFinishPoint(
          event
        ) != null
    ).length;

  return (
    <div className="analysis-pitch-wrapper">

      <h3>
        {title}
      </h3>

      <p className="heatmap-info">
        {positionedCount} von{" "}
        {events.length} Ereignissen
        mit Abschlussposition
      </p>

      <div className="football-pitch analysis-pitch heatmap-pitch">

        <div className="pitch-halfway-line" />
        <div className="pitch-center-circle" />
        <div className="penalty-area penalty-area-top" />
        <div className="penalty-area penalty-area-bottom" />
        <div className="goal-area goal-area-top" />
        <div className="goal-area goal-area-bottom" />

        <div
          className="heatmap-grid"
          style={{
            gridTemplateColumns:
              `repeat(${cols}, 1fr)`,

            gridTemplateRows:
              `repeat(${rows}, 1fr)`
          }}
        >

          {cells.map(
            cell => {
              const intensity =
                maxCount > 0
                  ? cell.count /
                    maxCount
                  : 0;

              return (
                <div
                  key={`${cell.row}-${cell.col}`}
                  className={`heatmap-cell ${
                    cell.count > 0
                      ? "heatmap-cell-active"
                      : ""
                  }`}
                  style={{
                    opacity:
                      cell.count > 0
                        ? 0.2 +
                          intensity *
                            0.65
                        : 0
                  }}
                  title={
                    cell.count > 0
                      ? `${cell.count} Abschlüsse (${cell.percentage} %)`
                      : ""
                  }
                >
                  {cell.count > 0 && (
                    <span>
                      {cell.count}
                    </span>
                  )}
                </div>
              );
            }
          )}

        </div>

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
            "Team Events:",
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

      try {
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
      } catch (
        error
      ) {
        console.error(
          "Events:",
          error
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
      setSaveStatus("");

      setAssistScreenPoint(
        null
      );

      setFinishScreenPoint(
        null
      );

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
     ANALYSIS
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


  const goalPhaseStats =
    useMemo(
      () =>
        createStats(
          teamGoals,
          "phase"
        ),
      [
        teamGoals
      ]
    );


  const concededPhaseStats =
    useMemo(
      () =>
        createStats(
          teamConceded,
          "phase"
        ),
      [
        teamConceded
      ]
    );


  const goalCreationStats =
    useMemo(
      () =>
        createStats(
          teamGoals,
          "creation_type"
        ),
      [
        teamGoals
      ]
    );


  const concededCreationStats =
    useMemo(
      () =>
        createStats(
          teamConceded,
          "creation_type"
        ),
      [
        teamConceded
      ]
    );


  const minuteBuckets =
    useMemo(
      () => {
        const buckets =
          selectedTeam ===
            "U15"
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

              goalPercentage:
                percentage(
                  goals,
                  teamGoals.length
                ),

              conceded,

              concededPercentage:
                percentage(
                  conceded,
                  teamConceded.length
                )
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
     MATCH HELPERS
     ===================================================== */

  const getMatchTitle =
    (
      match: MatchItem
    ) =>
      typeof match.opponent ===
        "string"
        ? match.opponent
        : "Unbekannter Gegner";


  const getMatchDate =
    (
      match: MatchItem
    ) => {
      const raw =
        typeof match.date ===
        "string"
          ? match.date
          : typeof match.match_date ===
              "string"
            ? match.match_date
            : "";

      if (
        !raw
      ) {
        return "";
      }

      const date =
        new Date(
          raw
        );

      if (
        Number.isNaN(
          date.getTime()
        )
      ) {
        return raw;
      }

      return new Intl.DateTimeFormat(
        "de-DE",
        {
          day:
            "2-digit",

          month:
            "2-digit",

          year:
            "numeric"
        }
      ).format(
        date
      );
    };


  const getCompetition =
    (
      match: MatchItem
    ) =>
      typeof match.competition ===
        "string"
        ? match.competition
        : typeof match.competition_type ===
            "string"
          ? match.competition_type
          : "";


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
     EVENT LIST
     ===================================================== */

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

                  setEventToDelete(
                    null
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


        {/* =================================================
            TEAM AUSWERTUNG
            ================================================= */}

        {viewMode ===
          "analysis" ? (

          <section className="team-analysis-section">

            <h2>
              {selectedTeam} – Team-Auswertung
            </h2>

            {loadingTeamEvents ? (

              <p>
                Auswertung wird geladen …
              </p>

            ) : (
              <>

                {/* KPI */}

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


                {/* ALLE EVENTS */}

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


                {/* HEATMAP */}

                <h2 className="analysis-heading">
                  Abschluss-Heatmap
                </h2>

                <div className="analysis-grid">

                  <FinishHeatmap
                    events={
                      teamGoals
                    }

                    title="Tore – Abschlusspositionen"
                  />

                  <FinishHeatmap
                    events={
                      teamConceded
                    }

                    title="Gegentore – Abschlusspositionen"
                  />

                </div>


                {/* PHASEN */}

                <h2 className="analysis-heading">
                  Spielphasen
                </h2>

                <div className="events-grid">

                  <StatCard
                    title="Tore nach Phase"
                    stats={
                      goalPhaseStats
                    }
                  />

                  <StatCard
                    title="Gegentore nach Phase"
                    stats={
                      concededPhaseStats
                    }
                  />

                </div>


                {/* ENTSTEHUNG */}

                <h2 className="analysis-heading">
                  Entstehungsarten
                </h2>

                <div className="events-grid">

                  <StatCard
                    title="Tore – Entstehung"
                    stats={
                      goalCreationStats
                    }
                  />

                  <StatCard
                    title="Gegentore – Entstehung"
                    stats={
                      concededCreationStats
                    }
                  />

                </div>


                {/* SPIELMINUTEN */}

                <h2 className="analysis-heading">
                  Tore nach Spielminute
                </h2>

                <div className="card minute-analysis-card">

                  <div className="minute-table minute-table-header">

                    <strong>
                      Zeitraum
                    </strong>

                    <strong>
                      Tore
                    </strong>

                    <strong>
                      Gegentore
                    </strong>

                  </div>

                  {minuteBuckets.map(
                    bucket => (

                      <div
                        key={
                          bucket.label
                        }

                        className="minute-table minute-row"
                      >

                        <strong>
                          {bucket.label}
                        </strong>

                        <span>
                          {bucket.goals}
                          {" "}
                          ({bucket.goalPercentage}%)
                        </span>

                        <span>
                          {bucket.conceded}
                          {" "}
                          ({bucket.concededPercentage}%)
                        </span>

                      </div>

                    )
                  )}

                </div>

              </>
            )}

          </section>

        ) : selectedMatch ? (

          /* =================================================
             EINZELSPIEL
             ================================================= */

          <section className="match-detail">

            <button
              onClick={() => {
                setSelectedMatch(
                  null
                );

                setEvents([]);

                setShowEventForm(
                  false
                );

                setEventToDelete(
                  null
                );
              }}
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

            <p>
              {getMatchDate(
                selectedMatch
              )}
            </p>

            {getCompetition(
              selectedMatch
            ) && (
              <p>
                {getCompetition(
                  selectedMatch
                )}
              </p>
            )}


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


            {/* DELETE */}

            {eventToDelete && (

              <div className="delete-confirm-box">

                <strong>
                  {eventToDelete.event_type ===
                  "Gegentor"
                    ? "Gegentor löschen?"
                    : "Tor löschen?"}
                </strong>

                <p>
                  {eventToDelete.minute != null
                    ? `${eventToDelete.minute}. Minute`
                    : "Minute unbekannt"}
                </p>

                <div className="form-actions">

                  <button
                    onClick={() =>
                      void executeDeleteEvent()
                    }

                    disabled={
                      deletingEvent
                    }
                  >
                    {deletingEvent
                      ? "Wird gelöscht …"
                      : "Ja, löschen"}
                  </button>

                  <button
                    onClick={() => {
                      setEventToDelete(
                        null
                      );

                      setDeleteStatus(
                        ""
                      );
                    }}
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


            {/* FORM */}

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

                {goals.length === 0
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

                {concededGoals.length === 0
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

          /* =================================================
             SPIELLISTE
             ================================================= */

          <section className="matches-section">

            <h2>
              {selectedTeam} Spiele
            </h2>

            <p className="status">
              {matchesStatus}
            </p>

            {loadingMatches ? (

              <p>
                Daten werden geladen …
              </p>

            ) : (

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

                    <p>
                      <strong>
                        Datum:
                      </strong>{" "}
                      {getMatchDate(
                        match
                      )}
                    </p>

                    {getCompetition(
                      match
                    ) && (
                      <p>
                        <strong>
                          Bewerb:
                        </strong>{" "}
                        {getCompetition(
                          match
                        )}
                      </p>
                    )}

                    <p>
                      Öffnen →
                    </p>

                  </div>

                )
              )

            )}

          </section>

        )}

      </main>

    </div>
  );
}

export default App;