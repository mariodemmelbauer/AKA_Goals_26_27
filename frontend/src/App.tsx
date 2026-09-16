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

function App() {
  const [status, setStatus] = useState("Teams wird initialisiert …");
  const [userName, setUserName] = useState("");
  const [tokenStatus, setTokenStatus] = useState("");

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

          const response = await fetch("/api/me", {
            method: "GET",
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: "application/json"
            }
          });

          const data = (await response.json()) as MeResponse;

          if (!response.ok) {
            console.error("API /api/me Fehler:", data);

            setTokenStatus(
              data.error
                ? `SSO Validierung fehlgeschlagen: ${data.error}`
                : `SSO Validierung fehlgeschlagen (${response.status})`
            );

            return;
          }

          if (!data.authenticated) {
            setTokenStatus(
              "Teams SSO konnte nicht bestätigt werden"
            );
            return;
          }

          const validatedUserName =
            data.user?.name ??
            data.user?.username ??
            contextUserName;

          if (validatedUserName) {
            setUserName(validatedUserName);
          }

          console.log("Serverseitig validierter Teams Benutzer:", data);

          setTokenStatus(
            "Teams SSO erfolgreich serverseitig validiert"
          );
        } catch (authError) {
          console.error("Teams SSO Fehler:", authError);

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
            Angemeldet als <strong>{userName}</strong>
          </p>
        )}

        <p className="status">{status}</p>

        {tokenStatus && (
          <p className="status">{tokenStatus}</p>
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
      </main>
    </div>
  );
}

export default App;