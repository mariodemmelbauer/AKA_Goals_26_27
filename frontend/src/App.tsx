import { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import "./App.css";


function App() {
  const [status, setStatus] = useState("Teams wird initialisiert …");
  const [userName, setUserName] = useState("");
  const [tokenStatus, setTokenStatus] = useState("");

  useEffect(() => {
    const initTeams = async () => {
      try {
        await microsoftTeams.app.initialize();

        const context = await microsoftTeams.app.getContext();

        setUserName(
          context.user?.displayName ??
          context.user?.userPrincipalName ??
          ""
        );

        setStatus("Microsoft Teams erkannt");

        try {
          const token = await microsoftTeams.authentication.getAuthToken();

          if (token) {
            setTokenStatus("Teams SSO Token erfolgreich empfangen");
            console.log("Teams SSO Token:", token);
          }
        } catch (authError) {
          console.error("SSO Fehler:", authError);
          setTokenStatus("Teams erkannt, SSO Token konnte aber nicht geladen werden");
        }
      } catch {
        setStatus("AKA Goals läuft aktuell außerhalb von Microsoft Teams");
      }
    };

    initTeams();
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
            <p>Goals & Analysis</p>
          </div>

          <div className="card">
            <h3>U16</h3>
            <p>Goals & Analysis</p>
          </div>

          <div className="card">
            <h3>U18</h3>
            <p>Goals & Analysis</p>
          </div>

          <div className="card">
            <h3>JWR</h3>
            <p>Goals & Analysis</p>
          </div>

          <div className="card">
            <h3>Profis</h3>
            <p>Goals & Analysis</p>
          </div>
        </div>
      </main>
    </div>
  );
}

export default App;