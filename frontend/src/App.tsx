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

  match_date?: string;

  opponent?: string;

  competition?: string | null;

  home_away?: string | null;

  created_by?: string | null;

  created_at?: string | null;

  [key: string]: unknown;
};

type MatchesResponse = {
  success?: boolean;
  matches?: MatchItem[];
  error?: string;
  details?: string;
};

type MatchApiResponse = {
  success?: boolean;
  match?: MatchItem;
  deletedId?: number | string;
  error?: string;
  details?: string;
};

type AssistPoint = {
  order: number;
  x: number;
  y: number;
  zone?: string | null;
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

  assist_points?: AssistPoint[] | null;

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
   ASSIST POINTS / LEGACY
   ===================================================== */

function getAssistPoints(
  event: GoalEvent
): AssistPoint[] {
  if (
    Array.isArray(
      event.assist_points
    ) &&
    event.assist_points.length >
      0
  ) {
    return event.assist_points
      .filter(
        point =>
          Number.isFinite(
            Number(
              point.x
            )
          ) &&
          Number.isFinite(
            Number(
              point.y
            )
          )
      )
      .map(
        (
          point,
          index
        ) => ({
          order:
            index +
            1,

          x:
            Number(
              point.x
            ),

          y:
            Number(
              point.y
            ),

          zone:
            point.zone ??
            null
        })
      )
      .slice(
        0,
        3
      );
  }


  /*
   * Alte Datensätze:
   * nur letzter Assistpunkt vorhanden.
   */

  if (
    event.assist_x != null &&
    event.assist_y != null
  ) {
    return [
      {
        order:
          1,

        x:
          Number(
            event.assist_x
          ),

        y:
          Number(
            event.assist_y
          ),

        zone:
          event.assist_zone ??
          null
      }
    ];
  }


  return [];
}


/* =====================================================
   ZONENLOGIK
   ===================================================== */

const FIRST_THIRD =
  100 / 3;

const SECOND_THIRD =
  (
    100 /
    3
  ) *
  2;

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

const BOX_LEFT_END =
  45;

const BOX_CENTER_END =
  55;


function getHorizontalLane(
  y: number
): string {
  if (
    y <
    20
  ) {
    return "linker Flügel";
  }

  if (
    y <
    40
  ) {
    return "linker Halbraum";
  }

  if (
    y <=
    60
  ) {
    return "Zentrum";
  }

  if (
    y <=
    80
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
    total <=
    0
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
    ) =>
      string |
      null |
      undefined
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

          {events.flatMap(
            (
              event,
              eventIndex
            ) => {
              const assists =
                getAssistPoints(
                  event
                );

              const finish =
                getFinishPoint(
                  event
                );


              if (
                !finish
              ) {
                return [];
              }


              const allPoints:
                DatabasePoint[] =
                [
                  ...assists.map(
                    point => ({
                      x:
                        point.x,

                      y:
                        point.y
                    })
                  ),

                  finish
                ];


              if (
                allPoints.length <
                2
              ) {
                return [];
              }


              return allPoints
                .slice(
                  0,
                  -1
                )
                .map(
                  (
                    point,
                    index
                  ) => {
                    const next =
                      allPoints[
                        index +
                        1
                      ];

                    const a =
                      databaseToScreen(
                        point.x,
                        point.y
                      );

                    const b =
                      databaseToScreen(
                        next.x,
                        next.y
                      );


                    return (
                      <line
                        key={`line-${event.id ?? eventIndex}-${index}`}
                        x1={a.x}
                        y1={a.y}
                        x2={b.x}
                        y2={b.y}
                        vectorEffect="non-scaling-stroke"
                      />
                    );
                  }
                );
            }
          )}

        </svg>


        {events.flatMap(
          (
            event,
            eventIndex
          ) =>
            getAssistPoints(
              event
            ).map(
              (
                assist,
                index
              ) => {
                const point =
                  databaseToScreen(
                    assist.x,
                    assist.y
                  );


                return (
                  <div
                    key={`assist-${event.id ?? eventIndex}-${index}`}

                    className="analysis-end-point"

                    style={{
                      left:
                        `${point.x}%`,

                      top:
                        `${point.y}%`,

                      background:
                        "#111",

                      color:
                        "#fff",

                      borderColor:
                        "#fff"
                    }}
                  >
                    {index +
                      1}
                  </div>
                );
              }
            )
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
                    {item.count}
                    {" "}
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
    useState(
      false
    );

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


  /* =====================================================
     MATCH STATES
     ===================================================== */

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
    useState(
      false
    );

  const [
    selectedMatch,
    setSelectedMatch
  ] =
    useState<MatchItem | null>(
      null
    );

  const [
    showMatchForm,
    setShowMatchForm
  ] =
    useState(
      false
    );

  const [
    newMatchDate,
    setNewMatchDate
  ] =
    useState("");

  const [
    newOpponent,
    setNewOpponent
  ] =
    useState("");

  const [
    newCompetition,
    setNewCompetition
  ] =
    useState(
      "Punktspiel"
    );

  const [
    newHomeAway,
    setNewHomeAway
  ] =
    useState(
      "Heim"
    );

  const [
    savingMatch,
    setSavingMatch
  ] =
    useState(
      false
    );

  const [
    matchSaveStatus,
    setMatchSaveStatus
  ] =
    useState("");

  const [
    matchToDelete,
    setMatchToDelete
  ] =
    useState<MatchItem | null>(
      null
    );

  const [
    deletingMatch,
    setDeletingMatch
  ] =
    useState(
      false
    );

  const [
    matchDeleteStatus,
    setMatchDeleteStatus
  ] =
    useState("");


  /* =====================================================
     EVENT STATES
     ===================================================== */

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
    useState(
      false
    );

  const [
    showEventForm,
    setShowEventForm
  ] =
    useState(
      false
    );

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


  /* =====================================================
     MULTI ASSIST STATES
     ===================================================== */

  const [
    actionCount,
    setActionCount
  ] =
    useState<
      1 |
      2 |
      3
    >(
      1
    );

  const [
    actionScreenPoints,
    setActionScreenPoints
  ] =
    useState<ScreenPoint[]>(
      []
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
    useState(
      false
    );

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
    useState(
      false
    );

  const [
    deleteStatus,
    setDeleteStatus
  ] =
    useState("");


  /* =====================================================
     MOVE EVENT STATES
     ===================================================== */

  const [
    eventToMove,
    setEventToMove
  ] =
    useState<GoalEvent | null>(
      null
    );

  const [
    moveTargetTeam,
    setMoveTargetTeam
  ] =
    useState<TeamName>(
      "U15"
    );

  const [
    moveMatches,
    setMoveMatches
  ] =
    useState<MatchItem[]>(
      []
    );

  const [
    moveTargetMatchId,
    setMoveTargetMatchId
  ] =
    useState("");

  const [
    loadingMoveMatches,
    setLoadingMoveMatches
  ] =
    useState(
      false
    );

  const [
    movingEvent,
    setMovingEvent
  ] =
    useState(
      false
    );

  const [
    moveStatus,
    setMoveStatus
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
     MATCHES LADEN
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
          loaded.length ===
            0
            ? `Keine ${team}-Spiele gefunden`
            : `${loaded.length} ${team}-Spiele geladen`
        );
      } catch (
        error
      ) {
        console.error(
          error
        );

        setMatchesStatus(
          "Fehler beim Laden der Spiele"
        );
      } finally {
        setLoadingMatches(
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
     SPIEL ANLEGEN
     ===================================================== */

  const saveMatch =
    async () => {
      if (
        !newMatchDate
      ) {
        setMatchSaveStatus(
          "Bitte Datum auswählen."
        );

        return;
      }


      if (
        !newOpponent.trim()
      ) {
        setMatchSaveStatus(
          "Bitte Gegner eingeben."
        );

        return;
      }


      setSavingMatch(
        true
      );

      setMatchSaveStatus(
        "Spiel wird gespeichert …"
      );


      try {
        const token =
          await getTeamsToken();


        const response =
          await fetch(
            "/api/matches",
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
                  team:
                    selectedTeam,

                  match_date:
                    newMatchDate,

                  opponent:
                    newOpponent.trim(),

                  competition:
                    newCompetition,

                  home_away:
                    newHomeAway
                })
            }
          );


        const data =
          (await response.json()) as MatchApiResponse;


        if (
          !response.ok
        ) {
          setMatchSaveStatus(
            data.details
              ? `${data.error} – ${data.details}`
              : data.error ??
                  "Spiel konnte nicht gespeichert werden"
          );

          return;
        }


        setShowMatchForm(
          false
        );

        setNewMatchDate("");

        setNewOpponent("");

        setNewCompetition(
          "Punktspiel"
        );

        setNewHomeAway(
          "Heim"
        );

        setMatchSaveStatus("");


        await loadMatches(
          selectedTeam
        );
      } catch (
        error
      ) {
        console.error(
          error
        );

        setMatchSaveStatus(
          "Fehler beim Speichern des Spiels"
        );
      } finally {
        setSavingMatch(
          false
        );
      }
    };


  /* =====================================================
     SPIEL LÖSCHEN
     ===================================================== */

  const deleteMatch =
    async () => {
      if (
        matchToDelete?.id == null
      ) {
        return;
      }


      setDeletingMatch(
        true
      );

      setMatchDeleteStatus(
        "Spiel wird gelöscht …"
      );


      try {
        const token =
          await getTeamsToken();


        const response =
          await fetch(
            `/api/matches?id=${encodeURIComponent(
              String(
                matchToDelete.id
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


        const data =
          (await response.json()) as MatchApiResponse;


        if (
          !response.ok
        ) {
          setMatchDeleteStatus(
            data.error ??
              "Spiel konnte nicht gelöscht werden"
          );

          return;
        }


        setMatchToDelete(
          null
        );

        setMatchDeleteStatus("");

        setSelectedMatch(
          null
        );

        setEvents([]);


        await loadMatches(
          selectedTeam
        );
      } finally {
        setDeletingMatch(
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
          response.ok
        ) {
          setTeamEvents(
            Array.isArray(
              data.events
            )
              ? data.events
              : []
          );
        }
      } finally {
        setLoadingTeamEvents(
          false
        );
      }
    };


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

      setShowMatchForm(
        false
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

      setShowMatchForm(
        false
      );

      void loadTeamEvents(
        selectedTeam
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
      setFinishTouch("");
      setSetPieceType("");
      setComment("");

      setActionCount(
        1
      );

      setActionScreenPoints(
        []
      );

      setFinishScreenPoint(
        null
      );

      setSaveStatus("");

      setShowEventForm(
        true
      );
    };


  /* =====================================================
     PITCH CLICK
     ===================================================== */

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
        actionScreenPoints.length <
        actionCount
      ) {
        setActionScreenPoints(
          [
            ...actionScreenPoints,
            point
          ]
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


      /*
       * Wenn bereits alles gesetzt wurde,
       * startet der nächste Klick eine neue Kette.
       */

      setActionScreenPoints(
        [
          point
        ]
      );

      setFinishScreenPoint(
        null
      );
    };


  /* =====================================================
     LIVE DATABASE POINTS
     ===================================================== */

  const actionDatabasePoints =
    useMemo(
      () =>
        actionScreenPoints.map(
          (
            point,
            index
          ) => {
            const databasePoint =
              screenToDatabase(
                point.x,
                point.y
              );

            return {
              order:
                index +
                1,

              x:
                databasePoint.x,

              y:
                databasePoint.y,

              zone:
                calculateZone(
                  databasePoint.x,
                  databasePoint.y
                )
            };
          }
        ),
      [
        actionScreenPoints
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
     EVENT SPEICHERN
     ===================================================== */

  const saveEvent =
    async () => {
      if (
        !selectedMatch ||
        selectedMatch.id == null
      ) {
        return;
      }


      if (
        actionDatabasePoints.length !==
        actionCount
      ) {
        setSaveStatus(
          `Bitte ${actionCount} Aktion${actionCount === 1 ? "" : "en"} vor dem Abschluss setzen.`
        );

        return;
      }


      if (
        !finishDatabasePoint
      ) {
        setSaveStatus(
          "Bitte Abschlussposition setzen."
        );

        return;
      }


      const lastAssist =
        actionDatabasePoints[
          actionDatabasePoints.length -
          1
        ];


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
                    minute ===
                    ""
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

                  assist_points:
                    actionDatabasePoints,

                  /*
                   * Letzter Assist zusätzlich in
                   * bestehenden Legacy-Spalten.
                   */

                  assist_x:
                    lastAssist.x,

                  assist_y:
                    lastAssist.y,

                  assist_zone:
                    lastAssist.zone,

                  finish_x:
                    finishDatabasePoint.x,

                  finish_y:
                    finishDatabasePoint.y,

                  finish_zone:
                    currentFinishZone,

                  finish_touch:
                    finishTouch,

                  set_piece_type:
                    setPieceType,

                  comment
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

        setSaveStatus("");
      } finally {
        setSavingEvent(
          false
        );
      }
    };


  /* =====================================================
     EVENT DELETE
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

          setDeleteStatus("");

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
     MOVE MATCHES LADEN
     ===================================================== */

  const loadMoveMatches =
    async (
      team: TeamName
    ) => {
      setLoadingMoveMatches(
        true
      );

      setMoveMatches([]);

      setMoveTargetMatchId("");

      setMoveStatus("");


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
          setMoveStatus(
            data.error ??
              "Ziel-Spiele konnten nicht geladen werden"
          );

          return;
        }


        const loaded =
          Array.isArray(
            data.matches
          )
            ? data.matches
            : [];


        setMoveMatches(
          loaded
        );


        if (
          loaded.length >
          0 &&
          loaded[0].id != null
        ) {
          setMoveTargetMatchId(
            String(
              loaded[0].id
            )
          );
        }
      } finally {
        setLoadingMoveMatches(
          false
        );
      }
    };


  const openMoveEvent =
    (
      event: GoalEvent
    ) => {
      setEventToMove(
        event
      );

      setMoveTargetTeam(
        selectedTeam
      );

      setMoveStatus("");

      void loadMoveMatches(
        selectedTeam
      );
    };


  /* =====================================================
     EVENT VERSCHIEBEN
     ===================================================== */

  const executeMoveEvent =
    async () => {
      if (
        eventToMove?.id == null
      ) {
        return;
      }


      if (
        !moveTargetMatchId
      ) {
        setMoveStatus(
          "Bitte Ziel-Spiel auswählen."
        );

        return;
      }


      setMovingEvent(
        true
      );

      setMoveStatus(
        "Ereignis wird verschoben …"
      );


      try {
        const token =
          await getTeamsToken();


        const response =
          await fetch(
            `/api/events?id=${encodeURIComponent(
              String(
                eventToMove.id
              )
            )}`,
            {
              method:
                "PATCH",

              headers: {
                Authorization:
                  `Bearer ${token}`,

                "Content-Type":
                  "application/json"
              },

              body:
                JSON.stringify({
                  team:
                    moveTargetTeam,

                  match_id:
                    moveTargetMatchId
                })
            }
          );


        const data =
          (await response.json()) as EventsResponse;


        if (
          !response.ok
        ) {
          setMoveStatus(
            data.details
              ? `${data.error} – ${data.details}`
              : data.error ??
                  "Verschieben fehlgeschlagen"
          );

          return;
        }


        setEventToMove(
          null
        );

        setMoveStatus("");


        if (
          selectedMatch
        ) {
          await loadEvents(
            selectedMatch
          );
        }


        await loadMatches(
          selectedTeam
        );
      } finally {
        setMovingEvent(
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


  const formatMatchDate =
    (
      raw:
        string |
        undefined
    ) => {
      if (
        !raw
      ) {
        return "";
      }


      const date =
        new Date(
          `${raw}T12:00:00`
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
     EVENT RENDER
     ===================================================== */

  const renderEvent =
    (
      event: GoalEvent
    ) => {
      const eventAssistPoints =
        getAssistPoints(
          event
        );


      return (
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


          <p>
            Aktionen vor Abschluss:{" "}
            {eventAssistPoints.length}
          </p>


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
              Letzte Assist-Zone:{" "}
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


          <div className="form-actions">

            <button
              onClick={() =>
                openMoveEvent(
                  event
                )
              }
            >
              Verschieben
            </button>


            <button
              className="delete-event-button"

              onClick={() => {
                setEventToDelete(
                  event
                );

                setDeleteStatus("");
              }}
            >
              Löschen
            </button>

          </div>

        </div>
      );
    };


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

                  setShowMatchForm(
                    false
                  );

                  setEventToDelete(
                    null
                  );

                  setEventToMove(
                    null
                  );

                  setMatchToDelete(
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

          <section>

            <h2>
              {selectedTeam}
              {" – Team-Auswertung"}
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

          /* =================================================
             MATCH DETAIL
             ================================================= */

          <section>

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

                setEventToMove(
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


            {selectedMatch.match_date && (
              <p>
                {formatMatchDate(
                  selectedMatch.match_date
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


            {/* EVENT VERSCHIEBEN */}

            {eventToMove && (

              <div className="delete-confirm-box">

                <strong>
                  {eventToMove.event_type}
                  {" verschieben"}
                </strong>


                <p>
                  Aktuell:
                  {" "}
                  {selectedTeam}
                  {" – "}
                  {getMatchTitle(
                    selectedMatch
                  )}
                </p>


                <div className="event-form">

                  <label>
                    Ziel-Team

                    <select
                      value={
                        moveTargetTeam
                      }

                      onChange={
                        event => {
                          const team =
                            event.target.value as TeamName;

                          setMoveTargetTeam(
                            team
                          );

                          void loadMoveMatches(
                            team
                          );
                        }
                      }
                    >

                      {TEAMS.map(
                        team => (

                          <option
                            key={
                              team
                            }

                            value={
                              team
                            }
                          >
                            {team}
                          </option>

                        )
                      )}

                    </select>
                  </label>


                  <label>
                    Ziel-Spiel

                    <select
                      value={
                        moveTargetMatchId
                      }

                      disabled={
                        loadingMoveMatches
                      }

                      onChange={
                        event =>
                          setMoveTargetMatchId(
                            event.target.value
                          )
                      }
                    >

                      {moveMatches.length ===
                        0 && (

                        <option value="">
                          Keine Spiele vorhanden
                        </option>

                      )}


                      {moveMatches.map(
                        match => (

                          <option
                            key={
                              String(
                                match.id
                              )
                            }

                            value={
                              String(
                                match.id
                              )
                            }
                          >
                            {formatMatchDate(
                              match.match_date
                            )}
                            {" – "}
                            {getMatchTitle(
                              match
                            )}
                          </option>

                        )
                      )}

                    </select>
                  </label>

                </div>


                {moveStatus && (
                  <p className="status">
                    {moveStatus}
                  </p>
                )}


                <div className="form-actions">

                  <button
                    onClick={() =>
                      void executeMoveEvent()
                    }

                    disabled={
                      movingEvent ||
                      !moveTargetMatchId
                    }
                  >
                    {movingEvent
                      ? "Wird verschoben …"
                      : "Ereignis verschieben"}
                  </button>


                  <button
                    onClick={() => {
                      setEventToMove(
                        null
                      );

                      setMoveStatus("");
                    }}
                  >
                    Abbrechen
                  </button>

                </div>

              </div>

            )}


            {/* EVENT DELETE */}

            {eventToDelete && (

              <div className="delete-confirm-box">

                <strong>
                  Ereignis wirklich löschen?
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
                    {deletingEvent
                      ? "Wird gelöscht …"
                      : "Ja, löschen"}
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


            {/* EVENT FORM */}

            {showEventForm && (

              <div className="card event-form-card">

                <h2>
                  {eventType}
                  {" erfassen"}
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
                        Assist / letzter Passgeber

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
                    Anzahl Aktionen vor dem Abschluss

                    <select
                      value={
                        actionCount
                      }

                      onChange={
                        event => {
                          const value =
                            Number(
                              event.target.value
                            ) as
                              1 |
                              2 |
                              3;

                          setActionCount(
                            value
                          );

                          setActionScreenPoints(
                            []
                          );

                          setFinishScreenPoint(
                            null
                          );
                        }
                      }
                    >

                      <option value="1">
                        1 Aktion / Assist
                      </option>

                      <option value="2">
                        2 Aktionen / Assists
                      </option>

                      <option value="3">
                        3 Aktionen / Assists
                      </option>

                    </select>
                  </label>


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

                      rows={
                        3
                      }

                      onChange={
                        event =>
                          setComment(
                            event.target.value
                          )
                      }
                    />
                  </label>


                  <p className="pitch-help">

                    {actionScreenPoints.length <
                    actionCount
                      ? `Klick ${actionScreenPoints.length + 1}: Aktion ${actionScreenPoints.length + 1} setzen`
                      : !finishScreenPoint
                        ? `Klick ${actionCount + 1}: Abschluss setzen`
                        : "Aktionskette vollständig gesetzt"}

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


                    <svg
                      className="pitch-line-layer"
                      viewBox="0 0 100 100"
                      preserveAspectRatio="none"
                    >

                      {[
                        ...actionScreenPoints,

                        ...(finishScreenPoint
                          ? [
                              finishScreenPoint
                            ]
                          : [])
                      ]
                        .slice(
                          0,
                          -1
                        )
                        .map(
                          (
                            point,
                            index
                          ) => {
                            const completePoints =
                              [
                                ...actionScreenPoints,

                                ...(finishScreenPoint
                                  ? [
                                      finishScreenPoint
                                    ]
                                  : [])
                              ];

                            const next =
                              completePoints[
                                index +
                                1
                              ];

                            if (
                              !next
                            ) {
                              return null;
                            }


                            return (
                              <line
                                key={`input-line-${index}`}
                                x1={
                                  point.x
                                }
                                y1={
                                  point.y
                                }
                                x2={
                                  next.x
                                }
                                y2={
                                  next.y
                                }
                                vectorEffect="non-scaling-stroke"
                              />
                            );
                          }
                        )}

                    </svg>


                    {actionScreenPoints.map(
                      (
                        point,
                        index
                      ) => (

                        <div
                          key={`action-${index}`}

                          className="pitch-point assist-point"

                          style={{
                            left:
                              `${point.x}%`,

                            top:
                              `${point.y}%`
                          }}
                        >
                          {index +
                            1}
                        </div>

                      )
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
                        T
                      </div>

                    )}

                  </div>


                  {actionDatabasePoints.length >
                    0 && (

                    <div className="card">

                      <strong>
                        Automatische Zonen
                      </strong>


                      {actionDatabasePoints.map(
                        point => (

                          <p
                            key={
                              point.order
                            }
                          >
                            Aktion{" "}
                            {point.order}:
                            {" "}
                            {point.zone}
                          </p>

                        )
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
                      setActionScreenPoints(
                        []
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


            {/* MATCH GRAPHICS */}

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


            {/* EVENT LIST */}

            <div className="events-grid">

              <div className="card">

                <h2>
                  Tore ({goals.length})
                </h2>


                {goals.length ===
                  0 ? (

                  <p>
                    Keine Tore erfasst.
                  </p>

                ) : (

                  goals.map(
                    renderEvent
                  )

                )}

              </div>


              <div className="card">

                <h2>
                  Gegentore ({concededGoals.length})
                </h2>


                {concededGoals.length ===
                  0 ? (

                  <p>
                    Keine Gegentore erfasst.
                  </p>

                ) : (

                  concededGoals.map(
                    renderEvent
                  )

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
              {selectedTeam}
              {" Spiele"}
            </h2>


            <div className="match-management-actions">

              <button
                onClick={() => {
                  setShowMatchForm(
                    true
                  );

                  setMatchSaveStatus("");

                  setMatchToDelete(
                    null
                  );
                }}
              >
                + Neues Spiel
              </button>

            </div>


            {showMatchForm && (

              <div className="card match-form-card">

                <h3>
                  Neues {selectedTeam}-Spiel
                </h3>


                <div className="event-form">

                  <label>
                    Datum

                    <input
                      type="date"

                      value={
                        newMatchDate
                      }

                      onChange={
                        event =>
                          setNewMatchDate(
                            event.target.value
                          )
                      }
                    />
                  </label>


                  <label>
                    Gegner

                    <input
                      value={
                        newOpponent
                      }

                      onChange={
                        event =>
                          setNewOpponent(
                            event.target.value
                          )
                      }
                    />
                  </label>


                  <label>
                    Bewerb

                    <select
                      value={
                        newCompetition
                      }

                      onChange={
                        event =>
                          setNewCompetition(
                            event.target.value
                          )
                      }
                    >

                      <option value="Punktspiel">
                        Punktspiel
                      </option>

                      <option value="Testspiel">
                        Testspiel
                      </option>

                    </select>
                  </label>


                  <label>
                    Heim / Auswärts

                    <select
                      value={
                        newHomeAway
                      }

                      onChange={
                        event =>
                          setNewHomeAway(
                            event.target.value
                          )
                      }
                    >

                      <option value="Heim">
                        Heim
                      </option>

                      <option value="Auswärts">
                        Auswärts
                      </option>

                    </select>
                  </label>

                </div>


                {matchSaveStatus && (
                  <p className="status">
                    {matchSaveStatus}
                  </p>
                )}


                <div className="form-actions">

                  <button
                    onClick={() =>
                      void saveMatch()
                    }

                    disabled={
                      savingMatch
                    }
                  >
                    {savingMatch
                      ? "Speichern …"
                      : "Spiel speichern"}
                  </button>


                  <button
                    onClick={() =>
                      setShowMatchForm(
                        false
                      )
                    }
                  >
                    Abbrechen
                  </button>

                </div>

              </div>

            )}


            {matchToDelete && (

              <div className="delete-confirm-box">

                <strong>
                  Spiel wirklich löschen?
                </strong>


                <p>
                  {getMatchTitle(
                    matchToDelete
                  )}
                </p>


                <p className="delete-warning">
                  Alle Tore und Gegentore dieses Spiels werden ebenfalls gelöscht.
                </p>


                <div className="form-actions">

                  <button
                    onClick={() =>
                      void deleteMatch()
                    }

                    disabled={
                      deletingMatch
                    }
                  >
                    {deletingMatch
                      ? "Wird gelöscht …"
                      : "Ja, Spiel löschen"}
                  </button>


                  <button
                    onClick={() =>
                      setMatchToDelete(
                        null
                      )
                    }
                  >
                    Abbrechen
                  </button>

                </div>


                {matchDeleteStatus && (
                  <p className="status">
                    {matchDeleteStatus}
                  </p>
                )}

              </div>

            )}


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
                    key={
                      String(
                        match.id ??
                        index
                      )
                    }

                    className="card match-card"
                  >

                    <div
                      className="match-card-content"

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


                      {match.match_date && (
                        <p>
                          <strong>
                            Datum:
                          </strong>
                          {" "}
                          {formatMatchDate(
                            match.match_date
                          )}
                        </p>
                      )}


                      {match.competition && (
                        <p>
                          <strong>
                            Bewerb:
                          </strong>
                          {" "}
                          {match.competition}
                        </p>
                      )}


                      {match.home_away && (
                        <p>
                          <strong>
                            Ort:
                          </strong>
                          {" "}
                          {match.home_away}
                        </p>
                      )}


                      <p>
                        Öffnen →
                      </p>

                    </div>


                    <button
                      className="delete-match-button"

                      onClick={() =>
                        setMatchToDelete(
                          match
                        )
                      }
                    >
                      Spiel löschen
                    </button>

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
