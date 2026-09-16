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

  /* =====================================================
     FRISCHEN TEAMS TOKEN HOLEN
     ===================================================== */

  const getTeamsToken = async (): Promise<string> => {
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
     SPIELE LADEN
     ===================================================== */

  const loadMatches = async (
    team: TeamName
  ) => {
    if (!teamsReady) {
      return;
    }

    setLoadingMatches(true);
    setMatches([]);
    setMatchesStatus(
      `${team}-Spiele werden geladen …`
    );

    try {
      /*
       * Wichtig:
       * Für diesen API Request holen wir
       * einen frischen Teams Token.
       */
      const token =
        await getTeamsToken();

      const response = await fetch(
        `/api/matches?team=${encodeURIComponent(
          team
        )}`,
        {
          method: "GET",
          headers: {
            Authorization:
              `Bearer ${token}`,
            Accept: "application/json"
          }
        }
      );

      let data: MatchesResponse;

      try {
        data =
          (await response.json()) as MatchesResponse;
      } catch {
        throw new Error(
          `Ungültige API-Antwort (${response.status})`
        );
      }

      if (!response.ok) {
        console.error(
          "Matches API Fehler:",
          data
        );

        setMatches([]);

        if (data.details) {
          console.error(
            "Supabase Details:",
            data.details
          );
        }

        setMatchesStatus(
          data.error
            ? `${data.error}${
                data.details
                  ? ` – ${data.details}`
                  : ""
              }`
            : `Spiele konnten nicht geladen werden (${response.status})`
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

      setMatches([]);

      setMatchesStatus(
        error instanceof Error
          ? error.message
          : `Fehler beim Laden der ${team}-Spiele`
      );
    } finally {
      setLoadingMatches(false);
    }
  };

  /* =====================================================
     TEAMS + SSO INITIALISIEREN
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

        setTokenStatus(
          "Teams SSO Token empfangen – Validierung läuft …"
        );

        const meResponse = await fetch(
          "/api/me",
          {
            method: "GET",
            headers: {
              Authorization:
                `Bearer ${token}`,
              Accept:
                "application/json"
            }
          }
        );

        const meData =
          (await meResponse.json()) as MeResponse;

        if (!meResponse.ok) {
          console.error(
            "/api/me Fehler:",
            meData
          );

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

        setTokenStatus(
          "Teams SSO erfolgreich serverseitig validiert"
        );

        /*
         * Erst jetzt API Aufrufe freigeben.
         */
        setTeamsReady(true);
      } catch (error) {
        console.error(
          "Teams Initialisierung:",
          error
        );

        setStatus(
          "AKA Goals läuft aktuell außerhalb von Microsoft Teams"
        );

        setTokenStatus("");
        setTeamsReady(false);
      }
    };

    void initTeams();
  }, []);

  /* =====================================================
     NACH ERFOLGREICHEM LOGIN U15 LADEN
     ===================================================== */

  useEffect(() => {
    if (!teamsReady) {
      return;
    }

    void loadMatches(
      selectedTeam
    );
  }, [teamsReady]);

  /* =====================================================
     TEAM AUSWÄHLEN
     ===================================================== */

  const selectTeam = (
    team: TeamName
  ) => {
    setSelectedTeam(team);

    void loadMatches(team);
  };

  /* =====================================================
     MATCH FORMATIERUNG
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
      const date =
        new Date(rawDate);

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
      typeof match.competition === "string" &&
      match.competition.trim()
    ) {
      return match.competition;
    }

    if (
      typeof match.competition_type === "string" &&
      match.competition_type.trim()
    ) {
      return match.competition_type;
    }

    return "";
  };

  /* =====================================================
     UI
     ===================================================== */

  return (
    <div className="app">
      <header>
        <div>
          <h1>
            AKA Goals
          </h1>

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
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (
                    event.key ===
                      "Enter" ||
                    event.key ===
                      " "
                  ) {
                    selectTeam(team);
                  }
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

        {/* SPIELE */}

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
                  marginTop:
                    "16px"
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