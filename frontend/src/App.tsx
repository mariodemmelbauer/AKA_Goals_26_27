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

type TeamSelection =
  | "Alle Teams"
  | TeamName;

type EventType =
  | "Tor"
  | "Gegentor";

type ViewMode =
  | "matches"
  | "analysis";

type PitchMode =
  | "chain"
  | "finish"
  | "assist";

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


const TEAM_SELECTIONS: TeamSelection[] = [
  "Alle Teams",
  ...TEAMS
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


/*
 * Muss exakt der DB-Constraint entsprechen:
 *
 * One Touch
 * Two Touch
 * >2 Touches
 */

const FINISH_TOUCHES = [
  "",
  "One Touch",
  "Two Touch",
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
   COORDINATES
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
   NORMALIZE EXISTING DATA TO TOP GOAL
   ===================================================== */

function flipPointToTopGoal(
  point: DatabasePoint
): DatabasePoint {
  return {
    x:
      100 -
      point.x,

    y:
      100 -
      point.y
  };
}


function shouldFlipEventToTopGoal(
  finish: DatabasePoint | null
): boolean {
  if (
    !finish
  ) {
    return false;
  }


  return finish.x <
    50;
}


function normalizePointForTopGoal(
  point: DatabasePoint,
  flip: boolean
): DatabasePoint {
  return flip
    ? flipPointToTopGoal(
        point
      )
    : point;
}


/* =====================================================
   ASSIST POINTS
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
   ZONES
   ===================================================== */

const FIRST_THIRD =
  100 /
  3;


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


  if (
    safeX >=
      SECOND_THIRD &&
    safeX <
      PENALTY_BOX_X &&
    safeY >=
      40 &&
    safeY <=
      60
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
   PENALTY
   ===================================================== */

function getPenaltyScreenPoint(
  eventType: EventType
): ScreenPoint {
  /*
   * Neue Eingabe soll immer Richtung oberes Tor erfolgen.
   *
   * Daher setzen wir auch bei Gegentor
   * den oberen Elfmeterpunkt.
   *
   * DB x = 89.52 entspricht oberem Angriffstor.
   */

  const distance =
    (
      11 /
      105
    ) *
    100;


  const databaseX =
    100 -
    distance;


  void eventType;


  return databaseToScreen(
    databaseX,
    50
  );
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
  values: string[]
): StatItem[] {
  const map =
    new Map<
      string,
      number
    >();


  values.forEach(
    value => {
      const key =
        value ||
        "Nicht angegeben";


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
            values.length
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


function zoneStats(
  events: GoalEvent[],
  mode:
    | "finish"
    | "assist"
): StatItem[] {
  const zones:
    string[] =
    [];


  events.forEach(
    event => {
      const finish =
        getFinishPoint(
          event
        );


      const flip =
        shouldFlipEventToTopGoal(
          finish
        );


      if (
        mode ===
        "finish"
      ) {
        if (
          finish
        ) {
          const displayPoint =
            normalizePointForTopGoal(
              finish,
              flip
            );


          zones.push(
            calculateZone(
              displayPoint.x,
              displayPoint.y
            )
          );
        }


        return;
      }


      getAssistPoints(
        event
      ).forEach(
        point => {
          const displayPoint =
            normalizePointForTopGoal(
              {
                x:
                  point.x,

                y:
                  point.y
              },

              flip
            );


          zones.push(
            calculateZone(
              displayPoint.x,
              displayPoint.y
            )
          );
        }
      );
    }
  );


  return createStats(
    zones
  );
}


/* =====================================================
   ZONE LABEL POSITIONS
   ===================================================== */

const ZONE_CENTERS:
  Record<
    string,
    ScreenPoint
  > =
{
  "Angriffsdrittel – linker Flügel":
    {
      x:
        10,

      y:
        25
    },

  "Angriffsdrittel – linker Halbraum":
    {
      x:
        30,

      y:
        25
    },

  "Zone 14 / Zentrum vor Box":
    {
      x:
        50,

      y:
        25
    },

  "Angriffsdrittel – rechter Halbraum":
    {
      x:
        70,

      y:
        25
    },

  "Angriffsdrittel – rechter Flügel":
    {
      x:
        90,

      y:
        25
    },

  "Box links":
    {
      x:
        33,

      y:
        9
    },

  "Box zentral":
    {
      x:
        50,

      y:
        9
    },

  "Box rechts":
    {
      x:
        67,

      y:
        9
    },

  "Mitteldrittel – linker Flügel":
    {
      x:
        10,

      y:
        50
    },

  "Mitteldrittel – linker Halbraum":
    {
      x:
        30,

      y:
        50
    },

  "Mitteldrittel – Zentrum":
    {
      x:
        50,

      y:
        50
    },

  "Mitteldrittel – rechter Halbraum":
    {
      x:
        70,

      y:
        50
    },

  "Mitteldrittel – rechter Flügel":
    {
      x:
        90,

      y:
        50
    },

  "Aufbaudrittel – linker Flügel":
    {
      x:
        10,

      y:
        83
    },

  "Aufbaudrittel – linker Halbraum":
    {
      x:
        30,

      y:
        83
    },

  "Aufbaudrittel – Zentrum":
    {
      x:
        50,

      y:
        83
    },

  "Aufbaudrittel – rechter Halbraum":
    {
      x:
        70,

      y:
        83
    },

  "Aufbaudrittel – rechter Flügel":
    {
      x:
        90,

      y:
        83
    }
};


/* =====================================================
   PITCH MARKINGS
   ===================================================== */

function PitchMarkings() {
  return (
    <>

      <svg
        className="pitch-markings"
        viewBox="0 0 100 100"
        preserveAspectRatio="none"
      >

        <rect
          x="0.5"
          y="0.5"
          width="99"
          height="99"
          className="pitch-white-line"
        />


        <line
          x1="0"
          y1="50"
          x2="100"
          y2="50"
          className="pitch-white-line"
        />


        <ellipse
          cx="50"
          cy="50"
          rx="13.5"
          ry="8.7"
          className="pitch-white-line"
        />


        <rect
          x="20.35"
          y="0"
          width="59.3"
          height="15.71"
          className="pitch-white-line"
        />


        <rect
          x="36.53"
          y="0"
          width="26.94"
          height="5.24"
          className="pitch-white-line"
        />


        <rect
          x="20.35"
          y="84.29"
          width="59.3"
          height="15.71"
          className="pitch-white-line"
        />


        <rect
          x="36.53"
          y="94.76"
          width="26.94"
          height="5.24"
          className="pitch-white-line"
        />


        <circle
          cx="50"
          cy="10.48"
          r="0.7"
          className="pitch-spot"
        />


        <circle
          cx="50"
          cy="89.52"
          r="0.7"
          className="pitch-spot"
        />


        <path
          d="M39.2 15.71 A13.5 8.7 0 0 0 60.8 15.71"
          className="pitch-white-line"
        />


        <path
          d="M39.2 84.29 A13.5 8.7 0 0 1 60.8 84.29"
          className="pitch-white-line"
        />

      </svg>


      <div
        className="pitch-goal pitch-goal-top"
      />


      <div
        className="pitch-goal pitch-goal-bottom"
      />

    </>
  );
}


/* =====================================================
   ZONE OVERLAY
   ===================================================== */

function ZoneOverlay({
  stats
}: {
  stats?: StatItem[];
}) {
  return (
    <div className="zone-overlay">

      <div className="zone-third-line zone-third-one" />

      <div className="zone-third-line zone-third-two" />

      <div className="zone-lane-line zone-lane-one" />

      <div className="zone-lane-line zone-lane-two" />

      <div className="zone-lane-line zone-lane-three" />

      <div className="zone-lane-line zone-lane-four" />

      <div className="zone14-highlight" />


      {stats?.map(
        item => {
          const center =
            ZONE_CENTERS[
              item.name
            ];


          if (
            !center
          ) {
            return null;
          }


          return (
            <div
              key={
                item.name
              }

              className="zone-percentage"

              title={
                item.name
              }

              style={{
                left:
                  `${center.x}%`,

                top:
                  `${center.y}%`
              }}
            >
              {item.percentage}%
            </div>
          );
        }
      )}

    </div>
  );
}


/* =====================================================
   ANALYSIS PITCH
   ===================================================== */

function AnalysisPitch({
  events,
  title,
  mode,
  zones = true
}: {
  events: GoalEvent[];
  title: string;
  mode: PitchMode;
  zones?: boolean;
}) {
  const stats =
    useMemo(
      () => {
        if (
          mode ===
          "assist"
        ) {
          return zoneStats(
            events,
            "assist"
          );
        }


        return zoneStats(
          events,
          "finish"
        );
      },
      [
        events,
        mode
      ]
    );


  return (
    <div className="analysis-pitch-wrapper">

      <h3>
        {title}
      </h3>


      <div className="football-pitch analysis-pitch">

        <PitchMarkings />


        {zones && (
          <ZoneOverlay
            stats={
              stats
            }
          />
        )}


        {mode ===
        "chain" && (

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


                const flip =
                  shouldFlipEventToTopGoal(
                    finish
                  );


                const databasePoints:
                  DatabasePoint[] =
                  [
                    ...assists.map(
                      point =>
                        normalizePointForTopGoal(
                          {
                            x:
                              point.x,

                            y:
                              point.y
                          },

                          flip
                        )
                    ),

                    normalizePointForTopGoal(
                      finish,
                      flip
                    )
                  ];


                return databasePoints
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
                        databasePoints[
                          index +
                          1
                        ];


                      const start =
                        databaseToScreen(
                          point.x,
                          point.y
                        );


                      const end =
                        databaseToScreen(
                          next.x,
                          next.y
                        );


                      return (
                        <line
                          key={`chain-${event.id ?? eventIndex}-${index}`}

                          x1={
                            start.x
                          }

                          y1={
                            start.y
                          }

                          x2={
                            end.x
                          }

                          y2={
                            end.y
                          }

                          vectorEffect="non-scaling-stroke"
                        />
                      );
                    }
                  );
              }
            )}

          </svg>

        )}


        {(mode ===
          "chain" ||
          mode ===
          "assist") &&

          events.flatMap(
            (
              event,
              eventIndex
            ) => {
              const finish =
                getFinishPoint(
                  event
                );


              const flip =
                shouldFlipEventToTopGoal(
                  finish
                );


              return getAssistPoints(
                event
              ).map(
                (
                  assist,
                  index
                ) => {
                  const displayPoint =
                    normalizePointForTopGoal(
                      {
                        x:
                          assist.x,

                        y:
                          assist.y
                      },

                      flip
                    );


                  const point =
                    databaseToScreen(
                      displayPoint.x,
                      displayPoint.y
                    );


                  return (
                    <div
                      key={`assist-${event.id ?? eventIndex}-${index}`}

                      className="analysis-assist-point"

                      style={{
                        left:
                          `${point.x}%`,

                        top:
                          `${point.y}%`
                      }}
                    >
                      {mode ===
                      "chain"
                        ? index +
                          1
                        : ""}
                    </div>
                  );
                }
              );
            }
          )}


        {(mode ===
          "chain" ||
          mode ===
          "finish") &&

          events.map(
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


              const flip =
                shouldFlipEventToTopGoal(
                  finish
                );


              const displayPoint =
                normalizePointForTopGoal(
                  finish,
                  flip
                );


              const point =
                databaseToScreen(
                  displayPoint.x,
                  displayPoint.y
                );


              return (
                <div
                  key={`finish-${event.id ?? index}`}

                  className="analysis-finish-point"

                  style={{
                    left:
                      `${point.x}%`,

                    top:
                      `${point.y}%`
                  }}
                >
                  {mode ===
                  "chain"
                    ? event.minute ??
                      ""
                    : ""}
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
    useState<TeamSelection>(
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
    const init =
      async () => {
        try {
          await microsoftTeams.app.initialize();


          const context =
            await microsoftTeams.app.getContext();


          setStatus(
            "Microsoft Teams erkannt"
          );


          const contextName =
            context.user?.displayName ??
            context.user?.userPrincipalName ??
            "";


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
            await response.json() as MeResponse;


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
            contextName
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


    void init();
  }, []);


  /* =====================================================
     LOAD MATCHES
     ===================================================== */

  const loadMatches =
    async (
      team: TeamSelection
    ) => {
      setLoadingMatches(
        true
      );


      setMatchesStatus(
        "Spiele werden geladen …"
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
          await response.json() as MatchesResponse;


        const loaded =
          response.ok &&
          Array.isArray(
            data.matches
          )
            ? data.matches
            : [];


        setMatches(
          loaded
        );


        setMatchesStatus(
          `${loaded.length} Spiele geladen`
        );
      } catch (
        error
      ) {
        console.error(
          error
        );


        setMatchesStatus(
          "Spiele konnten nicht geladen werden"
        );
      } finally {
        setLoadingMatches(
          false
        );
      }
    };


  /* =====================================================
     LOAD TEAM EVENTS
     ===================================================== */

  const loadTeamEvents =
    async (
      team: TeamSelection
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
          await response.json() as EventsResponse;


        setTeamEvents(
          response.ok &&
          Array.isArray(
            data.events
          )
            ? data.events
            : []
        );
      } catch (
        error
      ) {
        console.error(
          error
        );


        setTeamEvents(
          []
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
     TEAM SELECT
     ===================================================== */

  const selectTeam =
    (
      team: TeamSelection
    ) => {
      setSelectedTeam(
        team
      );


      setSelectedMatch(
        null
      );


      setEvents([]);


      setTeamEvents([]);


      setShowEventForm(
        false
      );


      setShowMatchForm(
        false
      );


      setEventToMove(
        null
      );


      setEventToDelete(
        null
      );


      setMatchToDelete(
        null
      );


      void loadMatches(
        team
      );


      if (
        team ===
        "Alle Teams"
      ) {
        setViewMode(
          "analysis"
        );


        void loadTeamEvents(
          team
        );


        return;
      }


      setViewMode(
        "matches"
      );
    };


  /* =====================================================
     LOAD EVENTS
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
        await response.json() as EventsResponse;


      setEvents(
        response.ok &&
        Array.isArray(
          data.events
        )
          ? data.events
          : []
      );
    };


  /* =====================================================
     MATCH CREATE
     ===================================================== */

  const saveMatch =
    async () => {
      if (
        selectedTeam ===
        "Alle Teams"
      ) {
        return;
      }


      if (
        !newMatchDate ||
        !newOpponent.trim()
      ) {
        setMatchSaveStatus(
          "Bitte Datum und Gegner eingeben."
        );


        return;
      }


      setSavingMatch(
        true
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
          await response.json() as MatchApiResponse;


        if (
          !response.ok
        ) {
          setMatchSaveStatus(
            data.error ??
            "Spiel konnte nicht gespeichert werden"
          );


          return;
        }


        setShowMatchForm(
          false
        );


        setNewMatchDate("");

        setNewOpponent("");

        setMatchSaveStatus("");


        await loadMatches(
          selectedTeam
        );
      } finally {
        setSavingMatch(
          false
        );
      }
    };


  /* =====================================================
     MATCH DELETE
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


      try {
        const token =
          await getTeamsToken();


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


        setMatchToDelete(
          null
        );


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
     OPEN EVENT
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
     SET PIECE
     ===================================================== */

  const handleSetPieceChange =
    (
      value: string
    ) => {
      setSetPieceType(
        value
      );


      if (
        value ===
        "Elfmeter"
      ) {
        setActionScreenPoints(
          []
        );


        setFinishScreenPoint(
          getPenaltyScreenPoint(
            eventType
          )
        );


        return;
      }


      setFinishScreenPoint(
        null
      );
    };


  /* =====================================================
     PITCH CLICK
     ===================================================== */

  const handlePitchClick =
    (
      event: MouseEvent<HTMLDivElement>
    ) => {
      if (
        setPieceType ===
        "Elfmeter"
      ) {
        return;
      }


      const rect =
        event.currentTarget.getBoundingClientRect();


      const point:
        ScreenPoint =
        {
          x:
            (
              (
                event.clientX -
                rect.left
              ) /
              rect.width
            ) *
            100,

          y:
            (
              (
                event.clientY -
                rect.top
              ) /
              rect.height
            ) *
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


      setActionScreenPoints(
        [
          point
        ]
      );


      setFinishScreenPoint(
        null
      );
    };


  const actionDatabasePoints =
    useMemo(
      () =>
        actionScreenPoints.map(
          (
            point,
            index
          ) => {
            const database =
              screenToDatabase(
                point.x,
                point.y
              );


            return {
              order:
                index +
                1,

              x:
                database.x,

              y:
                database.y,

              zone:
                calculateZone(
                  database.x,
                  database.y
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


  /* =====================================================
     SAVE EVENT
     ===================================================== */

  const saveEvent =
    async () => {
      if (
        selectedTeam ===
          "Alle Teams" ||
        !selectedMatch ||
        selectedMatch.id == null
      ) {
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


      const isPenalty =
        setPieceType ===
        "Elfmeter";


      if (
        !isPenalty &&
        actionDatabasePoints.length !==
          actionCount
      ) {
        setSaveStatus(
          "Bitte alle Aktionen vor dem Abschluss setzen."
        );


        return;
      }


      const lastAssist =
        actionDatabasePoints.length >
          0
          ? actionDatabasePoints[
              actionDatabasePoints.length -
              1
            ]
          : null;


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
                    minute
                      ? Number(
                          minute
                        )
                      : null,

                  scorer:
                    eventType ===
                    "Tor"
                      ? scorer
                      : null,

                  assister:
                    eventType ===
                      "Tor" &&
                    !isPenalty
                      ? assister
                      : null,

                  phase,

                  creation_type:
                    creationType,

                  assist_points:
                    isPenalty
                      ? []
                      : actionDatabasePoints,

                  assist_x:
                    isPenalty
                      ? null
                      : lastAssist?.x ??
                        null,

                  assist_y:
                    isPenalty
                      ? null
                      : lastAssist?.y ??
                        null,

                  assist_zone:
                    isPenalty
                      ? null
                      : lastAssist?.zone ??
                        null,

                  finish_x:
                    finishDatabasePoint.x,

                  finish_y:
                    finishDatabasePoint.y,

                  finish_zone:
                    calculateZone(
                      finishDatabasePoint.x,
                      finishDatabasePoint.y
                    ),

                  finish_touch:
                    finishTouch,

                  set_piece_type:
                    setPieceType,

                  comment
                })
            }
          );


        const data =
          await response.json() as EventsResponse;


        if (
          !response.ok
        ) {
          console.error(
            "Event speichern fehlgeschlagen:",
            data
          );


          setSaveStatus(
            data.details
              ? `${data.error ?? "Speichern fehlgeschlagen"} – ${data.details}`
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
     DELETE EVENT
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


        setEventToDelete(
          null
        );


        await loadEvents(
          selectedMatch
        );
      } finally {
        setDeletingEvent(
          false
        );
      }
    };


  /* =====================================================
     MOVE EVENT
     ===================================================== */

  const loadMoveMatches =
    async (
      team: TeamName
    ) => {
      setLoadingMoveMatches(
        true
      );


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
          await response.json() as MatchesResponse;


        const loaded =
          response.ok &&
          Array.isArray(
            data.matches
          )
            ? data.matches
            : [];


        setMoveMatches(
          loaded
        );


        setMoveTargetMatchId(
          loaded[0]?.id != null
            ? String(
                loaded[0].id
              )
            : ""
        );
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
      if (
        selectedTeam ===
        "Alle Teams"
      ) {
        return;
      }


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


  const executeMoveEvent =
    async () => {
      if (
        eventToMove?.id == null ||
        !moveTargetMatchId
      ) {
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
          await response.json() as EventsResponse;


        if (
          !response.ok
        ) {
          setMoveStatus(
            data.error ??
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
      } finally {
        setMovingEvent(
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


  const goalTouchStats =
    createStats(
      teamGoals.map(
        event =>
          event.finish_touch ||
          "Nicht angegeben"
      )
    );


  const concededTouchStats =
    createStats(
      teamConceded.map(
        event =>
          event.finish_touch ||
          "Nicht angegeben"
      )
    );


  const goalSetPieceStats =
    createStats(
      teamGoals
        .filter(
          event =>
            Boolean(
              event.set_piece_type
            )
        )
        .map(
          event =>
            event.set_piece_type ||
            ""
        )
    );


  const concededSetPieceStats =
    createStats(
      teamConceded
        .filter(
          event =>
            Boolean(
              event.set_piece_type
            )
        )
        .map(
          event =>
            event.set_piece_type ||
            ""
        )
    );


  const goalFinishZoneStats =
    zoneStats(
      teamGoals,
      "finish"
    );


  const concededFinishZoneStats =
    zoneStats(
      teamConceded,
      "finish"
    );


  const goalAssistZoneStats =
    zoneStats(
      teamGoals,
      "assist"
    );


  const concededAssistZoneStats =
    zoneStats(
      teamConceded,
      "assist"
    );


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
     HELPERS
     ===================================================== */

  const formatMatchDate =
    (
      value?: string
    ) => {
      if (
        !value
      ) {
        return "";
      }


      const date =
        new Date(
          `${value}T12:00:00`
        );


      if (
        Number.isNaN(
          date.getTime()
        )
      ) {
        return value;
      }


      return new Intl.DateTimeFormat(
        "de-DE"
      ).format(
        date
      );
    };


  const getMatchTitle =
    (
      match: MatchItem
    ) =>
      typeof match.opponent ===
      "string"
        ? match.opponent
        : "Unbekannter Gegner";


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
            onClick={() =>
              setEventToDelete(
                event
              )
            }
          >
            Löschen
          </button>

        </div>

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


        {/* TEAM CARDS */}

        <div className="cards">

          {TEAM_SELECTIONS.map(
            team => (

              <div
                key={
                  team
                }

                className={`card team-card ${
                  team ===
                  "Alle Teams"
                    ? "overall-team-card"
                    : ""
                } ${
                  selectedTeam ===
                  team
                    ? "active-team"
                    : ""
                }`}

                onClick={() =>
                  selectTeam(
                    team
                  )
                }
              >

                {team ===
                "Alle Teams" ? (

                  <>

                    <div className="overall-team-badge">
                      Gesamt
                    </div>


                    <h3>
                      Alle Teams
                    </h3>


                    <p className="team-card-main">
                      Gesamt-Auswertung
                    </p>


                    <p className="team-card-sub">
                      U15 · U16 · U18 · JWR · Profis
                    </p>

                  </>

                ) : (

                  <>

                    <h3>
                      {team}
                    </h3>


                    <p className="team-card-main">
                      Goals &amp; Analysis
                    </p>

                  </>

                )}

              </div>

            )
          )}

        </div>


        {selectedTeam !==
        "Alle Teams" && (

          <div className="view-switcher">

            <button
              className={
                viewMode ===
                "matches"
                  ? "view-button-active"
                  : ""
              }

              onClick={() => {
                setViewMode(
                  "matches"
                );


                setSelectedMatch(
                  null
                );
              }}
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

              onClick={() => {
                setViewMode(
                  "analysis"
                );


                setSelectedMatch(
                  null
                );


                void loadTeamEvents(
                  selectedTeam
                );
              }}
            >
              Team-Auswertung
            </button>

          </div>

        )}


        {viewMode ===
        "analysis" ? (

          <section>

            <h2>
              {selectedTeam}
              {" – Auswertung"}
            </h2>


            {selectedTeam ===
            "Alle Teams" && (

              <div className="overall-summary-hero">

                <div className="overall-summary-title">
                  Gesamtauswertung aller Teams
                </div>


                <div className="overall-summary-subtitle">
                  U15 · U16 · U18 · Junge Wikinger · Profis
                </div>

              </div>

            )}


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
                  Assist + Abschluss
                </h2>


                <div className="analysis-grid">

                  <AnalysisPitch
                    events={
                      teamGoals
                    }

                    title={`${selectedTeam} – Tore: Assist + Abschluss`}

                    mode="chain"
                  />


                  <AnalysisPitch
                    events={
                      teamConceded
                    }

                    title={`${selectedTeam} – Gegentore: Assist + Abschluss`}

                    mode="chain"
                  />

                </div>


                <h2 className="analysis-heading">
                  Abschlüsse
                </h2>


                <div className="analysis-grid">

                  <AnalysisPitch
                    events={
                      teamGoals
                    }

                    title="Tore – Abschlüsse"

                    mode="finish"
                  />


                  <AnalysisPitch
                    events={
                      teamConceded
                    }

                    title="Gegentore – Abschlüsse"

                    mode="finish"
                  />

                </div>


                <h2 className="analysis-heading">
                  Assists
                </h2>


                <div className="analysis-grid">

                  <AnalysisPitch
                    events={
                      teamGoals
                    }

                    title="Tore – Assists"

                    mode="assist"
                  />


                  <AnalysisPitch
                    events={
                      teamConceded
                    }

                    title="Gegentore – Assists"

                    mode="assist"
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
                  Assist-Zonen
                </h2>


                <div className="events-grid">

                  <StatCard
                    title="Tore – Assists"
                    stats={
                      goalAssistZoneStats
                    }
                  />


                  <StatCard
                    title="Gegentore – Assists"
                    stats={
                      concededAssistZoneStats
                    }
                  />

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
                    title="Tore"
                    stats={
                      goalSetPieceStats
                    }
                  />


                  <StatCard
                    title="Gegentore"
                    stats={
                      concededSetPieceStats
                    }
                  />

                </div>

              </>

            )}

          </section>

        ) : selectedMatch ? (

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


                setEventToMove(
                  null
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


            {eventToMove && (

              <div className="delete-confirm-box">

                <strong>
                  Ereignis verschieben
                </strong>


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
                    disabled={
                      movingEvent ||
                      !moveTargetMatchId
                    }

                    onClick={() =>
                      void executeMoveEvent()
                    }
                  >
                    {movingEvent
                      ? "Wird verschoben …"
                      : "Verschieben"}
                  </button>


                  <button
                    onClick={() =>
                      setEventToMove(
                        null
                      )
                    }
                  >
                    Abbrechen
                  </button>

                </div>

              </div>

            )}


            {eventToDelete && (

              <div className="delete-confirm-box">

                <strong>
                  Ereignis wirklich löschen?
                </strong>


                <div className="form-actions">

                  <button
                    disabled={
                      deletingEvent
                    }

                    onClick={() =>
                      void executeDeleteEvent()
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

              </div>

            )}


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


                      {setPieceType !==
                      "Elfmeter" && (

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

                      )}

                    </>

                  )}


                  {setPieceType !==
                  "Elfmeter" && (

                    <label>
                      Aktionen vor Abschluss

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
                          1 Aktion
                        </option>

                        <option value="2">
                          2 Aktionen
                        </option>

                        <option value="3">
                          3 Aktionen
                        </option>

                      </select>
                    </label>

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
                        value => (

                          <option
                            key={
                              value
                            }

                            value={
                              value
                            }
                          >
                            {value ||
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
                        value => (

                          <option
                            key={
                              value
                            }

                            value={
                              value
                            }
                          >
                            {value ||
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
                        value => (

                          <option
                            key={
                              value
                            }

                            value={
                              value
                            }
                          >
                            {value ||
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
                          handleSetPieceChange(
                            event.target.value
                          )
                      }
                    >

                      {SET_PIECE_TYPES.map(
                        value => (

                          <option
                            key={
                              value
                            }

                            value={
                              value
                            }
                          >
                            {value ||
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
                    />
                  </label>


                  <p className="pitch-help">

                    {setPieceType ===
                    "Elfmeter"
                      ? "Elfmeterpunkt automatisch oben gesetzt"
                      : actionScreenPoints.length <
                        actionCount
                        ? `Aktion ${actionScreenPoints.length + 1} setzen`
                        : !finishScreenPoint
                          ? "Abschluss oben setzen"
                          : "Aktionskette vollständig"}

                  </p>


                  <div
                    className="football-pitch capture-pitch"

                    onClick={
                      handlePitchClick
                    }
                  >

                    <PitchMarkings />


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
                            const all =
                              [
                                ...actionScreenPoints,

                                ...(finishScreenPoint
                                  ? [
                                      finishScreenPoint
                                    ]
                                  : [])
                              ];


                            const next =
                              all[
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
                                key={`capture-line-${index}`}

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
                          key={`capture-action-${index}`}

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


                  <div className="form-actions">

                    {setPieceType !==
                    "Elfmeter" && (

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

                    )}


                    <button
                      disabled={
                        savingEvent
                      }

                      onClick={() =>
                        void saveEvent()
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


                  {saveStatus && (
                    <p className="status">
                      {saveStatus}
                    </p>
                  )}

                </div>

              </div>

            )}


            <div className="analysis-grid">

              <AnalysisPitch
                events={
                  goals
                }

                title={`Tore (${goals.length})`}

                mode="chain"
              />


              <AnalysisPitch
                events={
                  concededGoals
                }

                title={`Gegentore (${concededGoals.length})`}

                mode="chain"
              />

            </div>


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
                  Neues Spiel
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
                    disabled={
                      savingMatch
                    }

                    onClick={() =>
                      void saveMatch()
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
                  Alle Tore und Gegentore werden ebenfalls gelöscht.
                </p>


                <div className="form-actions">

                  <button
                    disabled={
                      deletingMatch
                    }

                    onClick={() =>
                      void deleteMatch()
                    }
                  >
                    {deletingMatch
                      ? "Wird gelöscht …"
                      : "Ja, löschen"}
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

              </div>

            )}


            <p className="status">
              {matchesStatus}
            </p>


            {loadingMatches ? (

              <p>
                Spiele werden geladen …
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