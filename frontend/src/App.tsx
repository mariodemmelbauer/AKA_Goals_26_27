import { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import "./App.css";

type TeamName =
  | "U15"
  | "U16"
  | "U18"
  | "JWR"
  | "Profis";

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

  [key: string]: unknown;
};

type EventsResponse = {
  success?: boolean;
  events?: GoalEvent[];
  error?: string;
  details?: string;
};

const TEAMS: TeamName[] = [
  "U15",
  "U16",
  "U18",
  "JWR",
  "Profis"
];

function App() {
  const [status, setStatus] =
    useState("Teams wird initialisiert …");

  const [userName, setUserName] =
    useState("");

  const [tokenStatus, setTokenStatus] =
    useState("");

  const [teamsReady, setTeamsReady] =
    useState(false);

  const [selectedTeam, setSelectedTeam] =
    useState<TeamName>("U15");

  const [matches, setMatches] =
    useState<MatchItem[]>([]);

  const [matchesStatus, setMatchesStatus] =
    useState("");

  const [loadingMatches, setLoadingMatches] =
    useState(false);

  const [selectedMatch, setSelectedMatch] =
    useState<MatchItem | null>(null);

  const [events, setEvents] =
    useState<GoalEvent[]>([]);

  const [eventsStatus, setEventsStatus] =
    useState("");

  const [loadingEvents, setLoadingEvents] =
    useState(false);

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
     TEAMS / SSO
     ===================================================== */

  useEffect(() => {
    const initTeams = async () => {
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

        if (contextUserName) {
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

        if (!response.ok) {
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

        setTeamsReady(true);
      } catch (error) {
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
      setLoadingMatches(true);

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

        if (!response.ok) {
          setMatchesStatus(
            data.error ??
              "Spiele konnten nicht geladen werden"
          );

          return;
        }

        const loaded =
          Array.isArray(data.matches)
            ? data.matches
            : [];

        setMatches(loaded);

        setMatchesStatus(
          loaded.length === 0
            ? `Keine ${team}-Spiele gefunden`
            : `${loaded.length} ${team}-Spiele geladen`
        );
      } catch (error) {
        console.error(
          "Matches Fehler:",
          error
        );

        setMatchesStatus(
          "Fehler beim Laden der Spiele"
        );
      } finally {
        setLoadingMatches(false);
      }
    };

  useEffect(() => {
    if (!teamsReady) {
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
      if (match.id == null) {
        return;
      }

      setSelectedMatch(match);

      setLoadingEvents(true);

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
              String(match.id)
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

        if (!response.ok) {
          setEventsStatus(
            data.error ??
              "Events konnten nicht geladen werden"
          );

          return;
        }

        const loaded =
          Array.isArray(data.events)
            ? data.events
            : [];

        setEvents(loaded);

        setEventsStatus(
          loaded.length === 0
            ? "Für dieses Spiel sind noch keine Tore oder Gegentore erfasst."
            : `${loaded.length} Ereignisse geladen`
        );
      } catch (error) {
        console.error(
          "Events Fehler:",
          error
        );

        setEventsStatus(
          "Fehler beim Laden der Events"
        );
      } finally {
        setLoadingEvents(false);
      }
    };

  /* =====================================================
     TEAM AUSWAHL
     ===================================================== */

  const selectTeam =
    (
      team: TeamName
    ) => {
      setSelectedTeam(team);

      setSelectedMatch(null);
      setEvents([]);

      void loadMatches(team);
    };

  /* =====================================================
     HELPERS
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
          day: "2-digit",
          month: "2-digit",
          year: "numeric"
        }
      ).format(date);
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

  const goals =
    events.filter(
      (event) =>
        event.event_type === "Tor"
    );

  const concededGoals =
    events.filter(
      (event) =>
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
            (team) => (
              <div
                key={team}
                className={`card team-card ${
                  selectedTeam === team
                    ? "active-team"
                    : ""
                }`}
                onClick={() =>
                  selectTeam(team)
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

        {/* MATCH DETAIL */}

        {selectedMatch ? (
          <section
            style={{
              marginTop: "32px"
            }}
          >
            <button
              onClick={() => {
                setSelectedMatch(
                  null
                );

                setEvents([]);
              }}
              style={{
                marginBottom:
                  "20px"
              }}
            >
              ← Zurück zu den Spielen
            </button>

            <h2>
              {
                selectedTeam
              }{" "}
              –{" "}
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
                {
                  getCompetition(
                    selectedMatch
                  )
                }
              </p>
            )}

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
                  gap: "20px",
                  marginTop:
                    "25px"
                }}
              >
                {/* TORE */}

                <div className="card">
                  <h2>
                    Tore ({goals.length})
                  </h2>

                  {goals.length === 0 && (
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
                      </div>
                    )
                  )}
                </div>

                {/* GEGENTORE */}

                <div className="card">
                  <h2>
                    Gegentore (
                    {
                      concededGoals.length
                    }
                    )
                  </h2>

                  {concededGoals.length ===
                    0 && (
                    <p>
                      Keine Gegentore
                      erfasst.
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
                      </div>
                    )
                  )}
                </div>
              </div>
            )}
          </section>
        ) : (
          /* SPIELLISTE */

          <section
            style={{
              marginTop: "32px"
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
                        {
                          getCompetition(
                            match
                          )
                        }
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