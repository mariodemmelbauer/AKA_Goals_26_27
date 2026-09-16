import {
  useEffect,
  useState,
  type MouseEvent
} from "react";

import * as microsoftTeams
  from "@microsoft/teams-js";

import "./App.css";

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
              if (
                event.start_x == null ||
                event.start_y == null ||
                event.end_x == null ||
                event.end_y == null
              ) {
                return null;
              }

              return (
                <line
                  key={`line-${event.id ?? index}`}

                  x1={
                    Number(
                      event.start_x
                    )
                  }

                  y1={
                    Number(
                      event.start_y
                    )
                  }

                  x2={
                    Number(
                      event.end_x
                    )
                  }

                  y2={
                    Number(
                      event.end_y
                    )
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
            if (
              event.start_x == null ||
              event.start_y == null
            ) {
              return null;
            }

            return (
              <div
                key={`start-${event.id ?? index}`}

                className="analysis-start-point"

                style={{
                  left:
                    `${Number(
                      event.start_x
                    )}%`,

                  top:
                    `${Number(
                      event.start_y
                    )}%`
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
            if (
              event.end_x == null ||
              event.end_y == null
            ) {
              return null;
            }

            return (
              <div
                key={`end-${event.id ?? index}`}

                className="analysis-end-point"

                style={{
                  left:
                    `${Number(
                      event.end_x
                    )}%`,

                  top:
                    `${Number(
                      event.end_y
                    )}%`
                }}
              >
                {event.minute != null
                  ? event.minute
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
     TOKEN
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
          loaded.length === 0
            ? "Noch keine Ereignisse erfasst."
            : `${loaded.length} Ereignisse geladen`
        );
      } finally {
        setLoadingEvents(
          false
        );
      }
    };

  /* =====================================================
     DELETE EVENT
     ===================================================== */

  const deleteEvent =
    async (
      event: GoalEvent
    ) => {
      if (
        event.id == null ||
        !selectedMatch
      ) {
        return;
      }

      const label =
        event.event_type ===
        "Gegentor"
          ? "dieses Gegentor"
          : "dieses Tor";

      const confirmed =
        window.confirm(
          `Möchtest du ${label} wirklich löschen?`
        );

      if (
        !confirmed
      ) {
        return;
      }

      try {
        const token =
          await getTeamsToken();

        const response =
          await fetch(
            `/api/events?id=${encodeURIComponent(
              String(
                event.id
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
          window.alert(
            data.error ??
              "Event konnte nicht gelöscht werden."
          );

          return;
        }

        await loadEvents(
          selectedMatch
        );
      } catch (
        error
      ) {
        console.error(
          error
        );

        window.alert(
          "Fehler beim Löschen des Ereignisses."
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

      const point = {
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

      setAssistPoint(
        point
      );

      setFinishPoint(
        null
      );
    };

  /* =====================================================
     SAVE EVENT
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
        minute.trim() === ""
          ? null
          : Number(
              minute
            );

      if (
        parsedMinute !== null &&
        (
          !Number.isInteger(
            parsedMinute
          ) ||
          parsedMinute < 0 ||
          parsedMinute > 130
        )
      ) {
        setSaveStatus(
          "Bitte eine gültige Minute eingeben."
        );

        return;
      }

      if (
        !assistPoint
      ) {
        setSaveStatus(
          "Bitte die Entstehungsposition setzen."
        );

        return;
      }

      if (
        !finishPoint
      ) {
        setSaveStatus(
          "Bitte die Abschlussposition setzen."
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
                    parsedMinute,

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
            data.details
              ? `${data.error} – ${data.details}`
              : data.error ??
                  "Event konnte nicht gespeichert werden."
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
      } finally {
        setSavingEvent(
          false
        );
      }
    };

  /* =====================================================
     TEAM
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

      void loadMatches(
        team
      );
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
        new Date(
          raw
        );

      return Number.isNaN(
        date.getTime()
      )
        ? raw
        : new Intl.DateTimeFormat(
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
     EVENT LIST
     ===================================================== */

  const renderEvent =
    (
      event: GoalEvent,
      index: number
    ) => (
      <div
        key={String(
          event.id ??
            index
        )}
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
          onClick={() =>
            void deleteEvent(
              event
            )
          }
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
                  selectedTeam === team
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

        {selectedMatch ? (

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
                      onChange={(
                        event
                      ) =>
                        setMinute(
                          event.target.value
                        )
                      }
                    />
                  </label>

                  {eventType === "Tor" && (
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

                  <div className="pitch-section">

                    <h3>
                      Positionen
                    </h3>

                    <p className="pitch-help">
                      {!assistPoint
                        ? "1. Klick: Entstehung / Assist"
                        : !finishPoint
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

            <p className="status">
              {eventsStatus}
            </p>

            {loadingEvents ? (

              <p>
                Daten werden geladen …
              </p>

            ) : (
              <>

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

                    {goals.length === 0 ? (
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

                    {concededGoals.length === 0 ? (
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

              </>
            )}

          </section>

        ) : (

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