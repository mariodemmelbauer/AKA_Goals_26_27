import { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import "./App.css";

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

function App() {
  const [status, setStatus] = useState("Teams wird initialisiert …");
  const [userName, setUserName] = useState("");
  const [tokenStatus, setTokenStatus] = useState("");

  const [matches, setMatches] = useState<MatchItem[]>([]);
  const [matchesStatus, setMatchesStatus] = useState("");
  const [loadingMatches, setLoadingMatches] = useState(false);

  useEffect(() => {
    const initTeams = async () => {
      try {
        await microsoftTeams.app.initialize();

        const context = await microsoftTeams.app.getContext();

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

          /* =====================================================
             Benutzer serverseitig validieren
             ===================================================== */

          const meResponse = await fetch("/api/me", {
            method: "GET",
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: "application/json"
            }
          });

          const meData =
            (await meResponse.json()) as MeResponse;

          if (!meResponse.ok) {
            console.error(
              "API /api/me Fehler:",
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
            setUserName(validatedUserName);
          }

          setTokenStatus(
            "Teams SSO erfolgreich serverseitig validiert"
          );

          /* =====================================================
             U15-Spiele aus Supabase laden
             ===================================================== */

          setLoadingMatches(true);
          setMatchesStatus(
            "U15-Spiele werden geladen …"
          );

          try {
            const matchesResponse = await fetch(
              "/api/matches?team=U15",
              {
                method: "GET",
                headers: {
                  Authorization: `Bearer ${token}`,
                  Accept: "application/json"
                }
              }
            );

            const matchesData =
              (await matchesResponse.json()) as MatchesResponse;

            if (!matchesResponse.ok) {
              console.error(
                "API /api/matches Fehler:",
                matchesData
              );

              setMatches([]);

              setMatchesStatus(
                matchesData.error ??
                  `Spiele konnten nicht geladen werden (${matchesResponse.status})`
              );

              return;
            }

            const loadedMatches =
              Array.isArray(matchesData.matches)
                ? matchesData.matches
                : [];

            console.log(
              "U15 Matches:",
              loadedMatches
            );

            setMatches(loadedMatches);

            if (loadedMatches.length === 0) {
              setMatchesStatus(
                "Keine U15-Spiele gefunden"
              );
            } else {
              setMatchesStatus(
                `${loadedMatches.length} U15-Spiel${
                  loadedMatches.length === 1
                    ? ""
                    : "e"
                } geladen`
              );
            }
          } catch (matchesError) {
            console.error(
              "Fehler beim Laden der Spiele:",
              matchesError
            );

            setMatchesStatus(
              "Fehler beim Laden der U15-Spiele"
            );
          } finally {
            setLoadingMatches(false);
          }
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
        setMatches([]);
        setMatchesStatus("");
      }
    };

    void initTeams();
  }, []);

  const getMatchTitle = (
    match: MatchItem
  ): string => {
    const opponent =
      typeof match.opponent === "string"
        ? match.opponent
        : "Unbekannter Gegner";

    return opponent;
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
      return new Intl.DateTimeFormat(
        "de-DE",
        {
          day: "2-digit",
          month: "2-digit",
          year: "numeric"
        }
      ).format(new Date(rawDate));
    } catch {
      return rawDate;
    }
  };

  const getMatchCompetition = (
    match: MatchItem
  ): string => {
    if (
      typeof match.competition === "string"
    ) {
      return match.competition;
    }

    if (
      typeof match.competition_type === "string"
    ) {
      return match.competition_type;
    }

    return "";
  };

  return (
    <div className="app">
      <header>
        <div>
          <h1>AKA Goals</h1>
          <span>SV Oberbank Ried</span>
        </div>
      </header>

      <main>
        <h2>AKA Goals Dashboard</h2>

        {userName && (
          <p>
            Angemeldet als{" "}
            <strong>{userName}</strong>
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

        <div className="cards">
          <div className="card">
            <h3>U15</h3>
            <p>Goals &amp; Analysis</p>
          </div>

          <div className="card">
            <h3>U16</h3>
            <p>Goals &amp; Analysis</p>
          </div>

          <div className="card">
            <h3>U18</h3>
            <p>Goals &amp; Analysis</p>
          </div>

          <div className="card">
            <h3>JWR</h3>
            <p>Goals &amp; Analysis</p>
          </div>

          <div className="card">
            <h3>Profis</h3>
            <p>Goals &amp; Analysis</p>
          </div>
        </div>

        <section
          style={{
            marginTop: "32px"
          }}
        >
          <h2>U15 Spiele</h2>

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
                  (match, index) => {
                    const date =
                      getMatchDate(match);

                    const competition =
                      getMatchCompetition(
                        match
                      );

                    return (
                      <div
                        key={
                          String(
                            match.id ??
                              index
                          )
                        }
                        className="card"
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
                            {
                              String(
                                match.team
                              )
                            }
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