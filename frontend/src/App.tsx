import {
  useEffect,
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

  start_x?: number | null;
  start_y?: number | null;

  end_x?: number | null;
  end_y?: number | null;

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

type PitchPoint = {
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
   KOORDINATEN-KOMPATIBILITÄT
   ===================================================== */

function getStartX(
  event: GoalEvent
): number | null {
  if (
    event.start_x != null
  ) {
    return Number(
      event.start_x
    );
  }

  if (
    event.assist_x != null
  ) {
    return Number(
      event.assist_x
    );
  }

  return null;
}

function getStartY(
  event: GoalEvent
): number | null {
  if (
    event.start_y != null
  ) {
    return Number(
      event.start_y
    );
  }

  if (
    event.assist_y != null
  ) {
    return Number(
      event.assist_y
    );
  }

  return null;
}

function getFinishX(
  event: GoalEvent
): number | null {
  if (
    event.end_x != null
  ) {
    return Number(
      event.end_x
    );
  }

  if (
    event.finish_x != null
  ) {
    return Number(
      event.finish_x
    );
  }

  return null;
}

function getFinishY(
  event: GoalEvent
): number | null {
  if (
    event.end_y != null
  ) {
    return Number(
      event.end_y
    );
  }

  if (
    event.finish_y != null
  ) {
    return Number(
      event.finish_y
    );
  }

  return null;
}


/* =====================================================
   ANALYSE-SPIELFELD
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
              const startX =
                getStartX(
                  event
                );

              const startY =
                getStartY(
                  event
                );

              const finishX =
                getFinishX(
                  event
                );

              const finishY =
                getFinishY(
                  event
                );

              if (
                startX == null ||
                startY == null ||
                finishX == null ||
                finishY == null
              ) {
                return null;
              }

              return (
                <line
                  key={`line-${event.id ?? index}`}

                  x1={
                    startX
                  }

                  y1={
                    startY
                  }

                  x2={
                    finishX
                  }

                  y2={
                    finishY
                  }

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
            const startX =
              getStartX(
                event
              );

            const startY =
              getStartY(
                event
              );

            if (
              startX == null ||
              startY == null
            ) {
              return null;
            }

            return (
              <div
                key={`start-${event.id ?? index}`}

                className="analysis-start-point"

                style={{
                  left:
                    `${startX}%`,

                  top:
                    `${startY}%`
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
            const finishX =
              getFinishX(
                event
              );

            const finishY =
              getFinishY(
                event
              );

            if (
              finishX == null ||
              finishY == null
            ) {
              return null;
            }

            return (
              <div
                key={`finish-${event.id ?? index}`}

                className="analysis-end-point"

                style={{
                  left:
                    `${finishX}%`,

                  top:
                    `${finishY}%`
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
    events,
    setEvents
  ] =
    useState<GoalEvent[]>(
      []
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
    assistPoint,
    setAssistPoint
  ] =
    useState<PitchPoint | null>(
      null
    );

  const [
    finishPoint,
    setFinishPoint
  ] =
    useState<PitchPoint | null>(
      null
    );

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
     TEAMS INIT
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
          loaded.length ===
            0
            ? `Keine ${team}-Spiele gefunden`
            : `${loaded.length} ${team}-Spiele geladen`
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
     EVENTS
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
          !response.ok
        ) {
          console.error(
            "Events konnten nicht geladen werden:",
            data
          );

          return;
        }

        const loaded =
          Array.isArray(
            data.events
          )
            ? data.events
            : [];

        setEvents(
          loaded
        );
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

        const data =
          (await response.json()) as EventsResponse;

        if (
          !response.ok
        ) {
          setDeleteStatus(
            data.error ??
              "Event konnte nicht gelöscht werden"
          );

          return;
        }

        setEventToDelete(
          null
        );

        setDeleteStatus(
          ""
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

      setAssistPoint(
        null
      );

      setFinishPoint(
        null
      );

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
        PitchPoint =
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
        !assistPoint
      ) {
        setAssistPoint(
          point
        );

        return;
      }

      if (
        !finishPoint
      ) {
        setFinishPoint(
          point
        );

        return;
      }

      setAssistPoint(
        point
      );

      setFinishPoint(
        null
      );
    };


  /* =====================================================
     SAVE
     ===================================================== */

  const saveEvent =
    async () => {
      if (
        !selectedMatch ||
        selectedMatch.id == null ||
        !assistPoint ||
        !finishPoint
      ) {
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

                  start_x:
                    assistPoint.x,

                  start_y:
                    assistPoint.y,

                  end_x:
                    finishPoint.x,

                  end_y:
                    finishPoint.y
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

        setAssistPoint(
          null
        );

        setFinishPoint(
          null
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


        {/* TEAM AUSWAHL */}

        <div className="cards">

          {TEAMS.map(
            team => (
              <div
                key={team}

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


        {selectedMatch ? (
          <>

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


            {/* ACTIONS */}

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


            {/* DELETE CONFIRM */}

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


            {/* EVENT FORM */}

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


                    {assistPoint && (
                      <div
                        className="pitch-point assist-point"

                        style={{
                          left:
                            `${assistPoint.x}%`,

                          top:
                            `${assistPoint.y}%`
                        }}
                      >
                        1
                      </div>
                    )}


                    {finishPoint && (
                      <div
                        className="pitch-point finish-point"

                        style={{
                          left:
                            `${finishPoint.x}%`,

                          top:
                            `${finishPoint.y}%`
                        }}
                      >
                        2
                      </div>
                    )}


                    {assistPoint &&
                      finishPoint && (

                      <svg
                        className="pitch-line-layer"
                        viewBox="0 0 100 100"
                        preserveAspectRatio="none"
                      >
                        <line
                          x1={
                            assistPoint.x
                          }

                          y1={
                            assistPoint.y
                          }

                          x2={
                            finishPoint.x
                          }

                          y2={
                            finishPoint.y
                          }

                          vectorEffect="non-scaling-stroke"
                        />
                      </svg>

                    )}

                  </div>


                  <button
                    type="button"
                    onClick={() => {
                      setAssistPoint(
                        null
                      );

                      setFinishPoint(
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


            {/* ANALYSE */}

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


            {/* EVENT LISTEN */}

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
                    event => (
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
                    )
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
                    event => (
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
                    )
                  )
                )}

              </div>

            </div>

          </>
        ) : (

          /* SPIELLISTE */

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

                    {typeof match.match_date ===
                      "string" && (
                      <p>
                        {match.match_date}
                      </p>
                    )}

                    {typeof match.competition ===
                      "string" && (
                      <p>
                        {match.competition}
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