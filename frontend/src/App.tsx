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

  [key: string]: unknown;
};

type EventsResponse = {
  success?: boolean;

  events?: GoalEvent[];

  event?: GoalEvent;

  error?: string;

  details?: string;
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
   APP
   ===================================================== */

function App() {
  const [status, setStatus] =
    useState(
      "Teams wird initialisiert …"
    );

  const [userName, setUserName] =
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
    eventsStatus,
    setEventsStatus
  ] =
    useState("");

  const [
    loadingEvents,
    setLoadingEvents
  ] =
    useState(false);

  /* =====================================================
     EVENT FORM
     ===================================================== */

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
    savingEvent,
    setSavingEvent
  ] =
    useState(false);

  const [
    saveStatus,
    setSaveStatus
  ] =
    useState("");

  /* =====================================================
     PITCH
     ===================================================== */

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

  /* =====================================================
     TEAMS TOKEN
     ===================================================== */

  const getTeamsToken =
    async (): Promise<string> => {
      const token =
        await microsoftTeams.authentication.getAuthToken();

      if (!token) {
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
            context.user
              ?.displayName ??
            context.user
              ?.userPrincipalName ??
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
              data.user
                ?.username ??
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
            "Teams Init Fehler:",
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
      } catch (
        error
      ) {
        console.error(
          "Matches Fehler:",
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

      setLoadingEvents(
        true
      );

      setEvents([]);

      setEventsStatus(
        "Tore und Gegentore werden geladen …"
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
          setEventsStatus(
            data.error ??
              "Events konnten nicht geladen werden"
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

        setEventsStatus(
          loaded.length ===
            0
            ? "Für dieses Spiel sind noch keine Tore oder Gegentore erfasst."
            : `${loaded.length} Ereignisse geladen`
        );
      } catch (
        error
      ) {
        console.error(
          "Events Fehler:",
          error
        );

        setEventsStatus(
          "Fehler beim Laden der Events"
        );
      } finally {
        setLoadingEvents(
          false
        );
      }
    };

  /* =====================================================
     FORM ÖFFNEN
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

      const x =
        (
          (
            event.clientX -
            rect.left
          ) /
          rect.width
        ) *
        100;

      const y =
        (
          (
            event.clientY -
            rect.top
          ) /
          rect.height
        ) *
        100;

      const point: PitchPoint =
        {
          x:
            Math.round(
              x * 100
            ) / 100,

          y:
            Math.round(
              y * 100
            ) / 100
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

      /*
       * Wenn bereits beide Punkte
       * gesetzt sind, beginnt die
       * Eingabe wieder bei Punkt 1.
       */

      setAssistPoint(
        point
      );

      setFinishPoint(
        null
      );
    };

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

      const parsedMinute =
        minute.trim() ===
        ""
          ? null
          : Number(
              minute
            );

      if (
        parsedMinute !==
          null &&
        (
          !Number.isInteger(
            parsedMinute
          ) ||
          parsedMinute <
            0 ||
          parsedMinute >
            130
        )
      ) {
        setSaveStatus(
          "Bitte eine gültige Minute zwischen 0 und 130 eingeben."
        );

        return;
      }

      /* ---------- POSITIONEN ---------- */

      if (
        !assistPoint
      ) {
        setSaveStatus(
          "Bitte zuerst die Entstehungsposition auf dem Spielfeld setzen."
        );

        return;
      }

      if (
        !finishPoint
      ) {
        setSaveStatus(
          "Bitte anschließend die Abschlussposition auf dem Spielfeld setzen."
        );

        return;
      }

      setSavingEvent(
        true
      );

      setSaveStatus(
        "Ereignis wird gespeichert …"
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
                JSON.stringify(
                  {
                    match_id:
                      selectedMatch.id,

                    team:
                      selectedTeam,

                    event_type:
                      eventType,

                    minute:
                      parsedMinute,

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
                  }
                )
            }
          );

        const data =
          (await response.json()) as EventsResponse;

        if (
          !response.ok
        ) {
          console.error(
            "Event speichern Fehler:",
            data
          );

          setSaveStatus(
            data.details
              ? `${data.error} – ${data.details}`
              : data.error ??
                  "Event konnte nicht gespeichert werden"
          );

          return;
        }

        setSaveStatus(
          eventType ===
            "Tor"
            ? "Tor erfolgreich erfasst."
            : "Gegentor erfolgreich erfasst."
        );

        await loadEvents(
          selectedMatch
        );

        setTimeout(
          () => {
            setShowEventForm(
              false
            );

            setSaveStatus(
              ""
            );

            setAssistPoint(
              null
            );

            setFinishPoint(
              null
            );
          },
          1000
        );
      } catch (
        error
      ) {
        console.error(
          "Event speichern:",
          error
        );

        setSaveStatus(
          "Fehler beim Speichern des Ereignisses."
        );
      } finally {
        setSavingEvent(
          false
        );
      }
    };

  /* =====================================================
     TEAM SELECT
     ===================================================== */

  const selectTeam =
    (
      team: TeamName
    ) => {
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

      setAssistPoint(
        null
      );

      setFinishPoint(
        null
      );

      void loadMatches(
        team
      );
    };

  /* =====================================================
     MATCH HELPERS
     ===================================================== */

  const getMatchTitle =
    (
      match: MatchItem
    ) => {
      return typeof match.opponent ===
        "string"
        ? match.opponent
        : "Unbekannter Gegner";
    };

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

      if (!raw) {
        return "";
      }

      const date =
        new Date(raw);

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
    ) => {
      if (
        typeof match.competition ===
        "string"
      ) {
        return match.competition;
      }

      if (
        typeof match.competition_type ===
        "string"
      ) {
        return match.competition_type;
      }

      return "";
    };

  /* =====================================================
     EVENT FILTER
     ===================================================== */

  const goals =
    events.filter(
      (
        event
      ) =>
        event.event_type ===
        "Tor"
    );

  const concededGoals =
    events.filter(
      (
        event
      ) =>
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

        {/* =================================================
            TEAMS
            ================================================= */}

        <div className="cards">

          {TEAMS.map(
            (
              team
            ) => (

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

                onClick={() =>
                  selectTeam(
                    team
                  )
                }
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

        {/* =================================================
            MATCH DETAIL
            ================================================= */}

        {selectedMatch ? (

          <section
            style={{
              marginTop:
                "32px"
            }}
          >

            <button
              onClick={() => {
                setSelectedMatch(
                  null
                );

                setEvents([]);

                setShowEventForm(
                  false
                );

                setAssistPoint(
                  null
                );

                setFinishPoint(
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

            {/* =================================================
                ACTIONS
                ================================================= */}

            <div
              style={{
                display:
                  "flex",

                justifyContent:
                  "center",

                gap:
                  "12px",

                flexWrap:
                  "wrap",

                margin:
                  "24px 0"
              }}
            >

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

            {/* =================================================
                FORM
                ================================================= */}

            {showEventForm && (

              <div
                className="card"

                style={{
                  maxWidth:
                    "650px",

                  margin:
                    "0 auto 30px",

                  textAlign:
                    "left"
                }}
              >

                <h2>
                  {eventType} erfassen
                </h2>

                <div className="event-form">

                  {/* ---------- MINUTE ---------- */}

                  <label>

                    Minute

                    <input
                      type="number"

                      min="0"

                      max="130"

                      value={
                        minute
                      }

                      onChange={(
                        event
                      ) =>
                        setMinute(
                          event.target.value
                        )
                      }

                      placeholder="z. B. 35"
                    />

                  </label>

                  {/* ---------- TOR ---------- */}

                  {eventType ===
                    "Tor" && (
                    <>

                      <label>

                        Torschütze

                        <input
                          type="text"

                          value={
                            scorer
                          }

                          onChange={(
                            event
                          ) =>
                            setScorer(
                              event.target.value
                            )
                          }

                          placeholder="Name"
                        />

                      </label>

                      <label>

                        Assist

                        <input
                          type="text"

                          value={
                            assister
                          }

                          onChange={(
                            event
                          ) =>
                            setAssister(
                              event.target.value
                            )
                          }

                          placeholder="Name"
                        />

                      </label>

                    </>
                  )}

                  {/* ---------- PHASE ---------- */}

                  <label>

                    Phase

                    <select
                      value={
                        phase
                      }

                      onChange={(
                        event
                      ) =>
                        setPhase(
                          event.target.value
                        )
                      }
                    >

                      {PHASES.map(
                        (
                          item
                        ) => (

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

                  {/* ---------- CREATION ---------- */}

                  <label>

                    Entstehung

                    <select
                      value={
                        creationType
                      }

                      onChange={(
                        event
                      ) =>
                        setCreationType(
                          event.target.value
                        )
                      }
                    >

                      {CREATION_TYPES.map(
                        (
                          item
                        ) => (

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

                  {/* =================================================
                      PITCH
                      ================================================= */}

                  <div className="pitch-section">

                    <h3>
                      Positionen
                    </h3>

                    <p className="pitch-help">

                      {!assistPoint
                        ? "1. Klick: Entstehung / Assist"
                        : !finishPoint
                          ? "2. Klick: Abschluss"
                          : "Beide Positionen gesetzt – erneut klicken zum Neustart"}

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

                      {/* ---------- START ---------- */}

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

                      {/* ---------- END ---------- */}

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

                      {/* ---------- LINE ---------- */}

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

                    {/* ---------- COORDINATES ---------- */}

                    {(assistPoint ||
                      finishPoint) && (

                      <div className="pitch-values">

                        {assistPoint && (

                          <span>
                            1:{" "}
                            {assistPoint.x.toFixed(
                              1
                            )}{" "}
                            /{" "}
                            {assistPoint.y.toFixed(
                              1
                            )}
                          </span>

                        )}

                        {finishPoint && (

                          <span>
                            2:{" "}
                            {finishPoint.x.toFixed(
                              1
                            )}{" "}
                            /{" "}
                            {finishPoint.y.toFixed(
                              1
                            )}
                          </span>

                        )}

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

                    )}

                  </div>

                </div>

                {/* =================================================
                    SAVE STATUS
                    ================================================= */}

                {saveStatus && (
                  <p className="status">
                    {saveStatus}
                  </p>
                )}

                <div
                  style={{
                    display:
                      "flex",

                    gap:
                      "10px",

                    marginTop:
                      "20px"
                  }}
                >

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
                    onClick={() => {
                      setShowEventForm(
                        false
                      );

                      setSaveStatus(
                        ""
                      );

                      setAssistPoint(
                        null
                      );

                      setFinishPoint(
                        null
                      );
                    }}
                  >
                    Abbrechen
                  </button>

                </div>

              </div>

            )}

            {/* =================================================
                EVENTS
                ================================================= */}

            <p className="status">
              {eventsStatus}
            </p>

            {loadingEvents && (
              <p>
                Daten werden geladen …
              </p>
            )}

            {!loadingEvents && (

              <div
                style={{
                  display:
                    "grid",

                  gridTemplateColumns:
                    "repeat(auto-fit, minmax(300px, 1fr))",

                  gap:
                    "20px",

                  marginTop:
                    "25px"
                }}
              >

                {/* =================================================
                    TORE
                    ================================================= */}

                <div className="card">

                  <h2>
                    Tore ({goals.length})
                  </h2>

                  {goals.length ===
                    0 && (
                    <p>
                      Keine Tore erfasst.
                    </p>
                  )}

                  {goals.map(
                    (
                      event,
                      index
                    ) => (

                      <div
                        key={String(
                          event.id ??
                            index
                        )}

                        style={{
                          borderTop:
                            "1px solid #ddd",

                          padding:
                            "12px 0"
                        }}
                      >

                        <strong>
                          {event.minute !=
                          null
                            ? `${event.minute}. Minute`
                            : "Minute unbekannt"}
                        </strong>

                        {event.scorer && (
                          <p>
                            Torschütze:{" "}
                            {
                              event.scorer
                            }
                          </p>
                        )}

                        {event.assister && (
                          <p>
                            Assist:{" "}
                            {
                              event.assister
                            }
                          </p>
                        )}

                        {event.phase && (
                          <p>
                            Phase:{" "}
                            {
                              event.phase
                            }
                          </p>
                        )}

                        {event.creation_type && (
                          <p>
                            Entstehung:{" "}
                            {
                              event.creation_type
                            }
                          </p>
                        )}

                        {event.start_x !=
                          null &&
                          event.start_y !=
                            null && (

                          <p>
                            Start:{" "}
                            {Number(
                              event.start_x
                            ).toFixed(
                              1
                            )}{" "}
                            /{" "}
                            {Number(
                              event.start_y
                            ).toFixed(
                              1
                            )}
                          </p>

                        )}

                        {event.end_x !=
                          null &&
                          event.end_y !=
                            null && (

                          <p>
                            Abschluss:{" "}
                            {Number(
                              event.end_x
                            ).toFixed(
                              1
                            )}{" "}
                            /{" "}
                            {Number(
                              event.end_y
                            ).toFixed(
                              1
                            )}
                          </p>

                        )}

                      </div>

                    )
                  )}

                </div>

                {/* =================================================
                    GEGENTORE
                    ================================================= */}

                <div className="card">

                  <h2>
                    Gegentore (
                    {concededGoals.length}
                    )
                  </h2>

                  {concededGoals.length ===
                    0 && (
                    <p>
                      Keine Gegentore erfasst.
                    </p>
                  )}

                  {concededGoals.map(
                    (
                      event,
                      index
                    ) => (

                      <div
                        key={String(
                          event.id ??
                            index
                        )}

                        style={{
                          borderTop:
                            "1px solid #ddd",

                          padding:
                            "12px 0"
                        }}
                      >

                        <strong>
                          {event.minute !=
                          null
                            ? `${event.minute}. Minute`
                            : "Minute unbekannt"}
                        </strong>

                        {event.phase && (
                          <p>
                            Phase:{" "}
                            {
                              event.phase
                            }
                          </p>
                        )}

                        {event.creation_type && (
                          <p>
                            Entstehung:{" "}
                            {
                              event.creation_type
                            }
                          </p>
                        )}

                        {event.start_x !=
                          null &&
                          event.start_y !=
                            null && (

                          <p>
                            Start:{" "}
                            {Number(
                              event.start_x
                            ).toFixed(
                              1
                            )}{" "}
                            /{" "}
                            {Number(
                              event.start_y
                            ).toFixed(
                              1
                            )}
                          </p>

                        )}

                        {event.end_x !=
                          null &&
                          event.end_y !=
                            null && (

                          <p>
                            Abschluss:{" "}
                            {Number(
                              event.end_x
                            ).toFixed(
                              1
                            )}{" "}
                            /{" "}
                            {Number(
                              event.end_y
                            ).toFixed(
                              1
                            )}
                          </p>

                        )}

                      </div>

                    )
                  )}

                </div>

              </div>

            )}

          </section>

        ) : (

          /* =================================================
             SPIELE
             ================================================= */

          <section
            style={{
              marginTop:
                "32px"
            }}
          >

            <h2>
              {selectedTeam} Spiele
            </h2>

            <p className="status">
              {matchesStatus}
            </p>

            {loadingMatches && (
              <p>
                Daten werden geladen …
              </p>
            )}

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

                    style={{
                      cursor:
                        "pointer",

                      marginBottom:
                        "12px"
                    }}
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
              )}

          </section>

        )}

      </main>

    </div>
  );
}

export default App;