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
  objectId?: string | null;
  tenantId?: string | null;
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
  home_away?: string;
  [key: string]: unknown;
};

type MatchesResponse = {
  success?: boolean;
  matches?: MatchItem[];
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
  const [status, setStatus] = useState(
    "Teams wird initialisiert …"
  );

  const [userName, setUserName] =
    useState("");

  const [tokenStatus, setTokenStatus] =
    useState("");

  const [authToken, setAuthToken] =
    useState("");

  const [selectedTeam, setSelectedTeam] =
    useState<TeamName>("U15");

  const [matches, setMatches] =
    useState<MatchItem[]>([]);

  const [matchesStatus, setMatchesStatus] =
    useState("");

  const [loadingMatches, setLoadingMatches] =
    useState(false);

  /* =====================================================
     TEAMS / SSO INITIALISIERUNG
     ===================================================== */

  useEffect(() => {
    const initTeams = async () => {
      try {
        await microsoftTeams.app.initialize();

        const context =
          await microsoftTeams.app.getContext();

        setStatus("Microsoft Teams erkannt");

        const contextUserName =
          context.user?.displayName ??
          context.user?.userPrincipalName ??
          "";

        if (contextUserName) {
          setUserName(contextUserName);
        }

        try {
          const token =
            await microsoftTeams.authentication.getAuthToken();

          if (!token) {
            setTokenStatus(
              "Teams erkannt, aber kein SSO Token empfangen"
            );
            return;
          }

          setTokenStatus(
            "Teams SSO Token empfangen – Validierung läuft …"
          );

          const meResponse = await fetch(
            "/api/me",
            {
              method: "GET",
              headers: {
                Authorization: `Bearer ${token}`,
                Accept: "application/json"
              }
            }
          );

          const meData =
            (await meResponse.json()) as MeResponse;

          if (!meResponse.ok) {
            setTokenStatus(
              meData.error
                ? `SSO Validierung fehlgeschlagen: ${meData.error}`
                : `SSO Validierung fehlgeschlagen (${meResponse.status})`
            );

            return;
          }

          if (!meData.authenticated) {
            setTokenStatus(
              "Teams SSO konnte nicht bestätigt werden"
            );
            return;
          }

          const validatedUserName =
            meData.user?.name ??
            meData.user?.username ??
            contextUserName;

          if (validatedUserName) {
            setUserName(
              validatedUserName
            );
          }

          setAuthToken(token);

          setTokenStatus(
            "Teams SSO erfolgreich serverseitig validiert"
          );
        } catch (authError) {
          console.error(
            "Teams SSO Fehler:",
            authError
          );

          setTokenStatus(
            "Teams erkannt, SSO konnte aber nicht validiert werden"
          );
        }
      } catch (teamsError) {
        console.log(
          "AKA Goals läuft außerhalb von Microsoft Teams:",
          teamsError
        );

        setStatus(
          "AKA Goals läuft aktuell außerhalb von Microsoft Teams"
        );

        setTokenStatus("");
      }
    };

    void initTeams();
  }, []);

  /* =====================================================
     SPIELE LADEN
     ===================================================== */

  const loadMatches = async (
    team: TeamName
  ) => {
    if (!authToken) {
      return;
    }

    setLoadingMatches(true);
    setMatches([]);
    setMatchesStatus(
      `${team}-Spiele werden geladen …`
    );

    try {
      const response = await fetch(
        `/api/matches?team=${encodeURIComponent(
          team
        )}`,
        {
          method: "GET",
          headers: {
            Authorization:
              `Bearer ${authToken}`,
            Accept: "application/json"
          }
        }
      );

      const data =
        (await response.json()) as MatchesResponse;

      if (!response.ok) {
        console.error(
          "Matches API Fehler:",
          data
        );

        setMatches([]);

        setMatchesStatus(
          data.error ??
            `Spiele konnten nicht geladen werden (${response.status})`
        );

        return;
      }

      const loadedMatches =
        Array.isArray(data.matches)
          ? data.matches
          : [];

      setMatches(loadedMatches);

      if (loadedMatches.length === 0) {
        setMatchesStatus(
          `Keine ${team}-Spiele gefunden`
        );
      } else {
        setMatchesStatus(
          `${loadedMatches.length} ${team}-Spiel${
            loadedMatches.length === 1
              ? ""
              : "e"
          } geladen`
        );
      }
    } catch (error) {
      console.error(
        "Fehler beim Laden der Spiele:",
        error
      );

      setMatchesStatus(
        `Fehler beim Laden der ${team}-Spiele`
      );
    } finally {
      setLoadingMatches(false);
    }
  };

  /* =====================================================
     NACH SSO INITIAL U15 LADEN
     ===================================================== */

  useEffect(() => {
    if (!authToken) {
      return;
    }

    void loadMatches(
      selectedTeam
    );
  }, [authToken]);

  /* =====================================================
     TEAM WECHSELN
     ===================================================== */

  const selectTeam = (
    team: TeamName
  ) => {
    if (
      team === selectedTeam &&
      matches.length > 0
    ) {
      return;
    }

    setSelectedTeam(team);

    void loadMatches(team);
  };

  /* =====================================================
     MATCH HELPERS
     ===================================================== */

  const getMatchTitle = (
    match: MatchItem
  ): string => {
    if (
      typeof match.opponent === "string" &&
      match.opponent.trim()
    ) {
      return match.opponent;
    }

    return "Unbekannter Gegner";
  };

  const getMatchDate = (
    match: MatchItem
  ): string => {
    const rawDate =
      typeof match.date === "string"
        ? match.date
        : typeof match.match_date === "string"
          ? match.match_date
          : "";

    if (!rawDate) {
      return "";
    }

    try {
      const date = new Date(rawDate);

      if (
        Number.isNaN(
          date.getTime()
        )
      ) {
        return rawDate;
      }

      return new Intl.DateTimeFormat(
        "de-DE",
        {
          day: "2-digit",
          month: "2-digit",
          year: "numeric"
        }
      ).format(date);
    } catch {
      return rawDate;
    }
  };

  const getCompetition = (
    match: MatchItem
  ): string => {
    if (
      typeof match.competition ===
        "string" &&
      match.competition.trim()
    ) {
      return match.competition;
    }

    if (
      typeof match.competition_type ===
        "string" &&
      match.competition_type.trim()
    ) {
      return match.competition_type;
    }

    return "";
  };

  /* =====================================================
     RENDER
     ===================================================== */

  return (
    <div className="app">
      <header>
        <div>
          <h1>AKA Goals</h1>

          <span>
            SV Oberbank Ried
          </span>
        </div>
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

        {tokenStatus && (
          <p className="status">
            {tokenStatus}
          </p>
        )}

        {/* ================================
            TEAM AUSWAHL
            ================================ */}

        <div className="cards">
          {TEAMS.map((team) => (
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
              role="button"
              tabIndex={0}
              onKeyDown={(event) => {
                if (
                  event.key ===
                    "Enter" ||
                  event.key === " "
                ) {
                  selectTeam(team);
                }
              }}
            >
              <h3>{team}</h3>

              <p>
                Goals &amp; Analysis
              </p>
            </div>
          ))}
        </div>

        {/* ================================
            SPIELLISTE
            ================================ */}

        <section
          style={{
            marginTop: "32px"
          }}
        >
          <h2>
            {selectedTeam} Spiele
          </h2>

          {matchesStatus && (
            <p className="status">
              {matchesStatus}
            </p>
          )}

          {loadingMatches && (
            <p>
              Daten werden geladen …
            </p>
          )}

          {!loadingMatches &&
            matches.length > 0 && (
              <div
                style={{
                  display: "grid",
                  gap: "12px",
                  marginTop: "16px"
                }}
              >
                {matches.map(
                  (
                    match,
                    index
                  ) => {
                    const date =
                      getMatchDate(
                        match
                      );

                    const competition =
                      getCompetition(
                        match
                      );

                    return (
                      <div
                        key={String(
                          match.id ??
                            index
                        )}
                        className="card match-card"
                      >
                        <h3>
                          {getMatchTitle(
                            match
                          )}
                        </h3>

                        {date && (
                          <p>
                            <strong>
                              Datum:
                            </strong>{" "}
                            {date}
                          </p>
                        )}

                        {competition && (
                          <p>
                            <strong>
                              Bewerb:
                            </strong>{" "}
                            {
                              competition
                            }
                          </p>
                        )}

                        {match.team && (
                          <p>
                            <strong>
                              Team:
                            </strong>{" "}
                            {String(
                              match.team
                            )}
                          </p>
                        )}
                      </div>
                    );
                  }
                )}
              </div>
            )}
        </section>
      </main>
    </div>
  );
}

export default App;