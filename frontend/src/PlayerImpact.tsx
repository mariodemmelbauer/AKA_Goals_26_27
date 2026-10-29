import { useMemo, useState } from "react";

type Goal = [minute: number, side: "for" | "against"];
type Substitution = [minute: number, off: string, on: string];
type Fixture = {
  date: string;
  opponent: string;
  result: string;
  goals: Goal[];
  starters: string[];
  substitutions: Substitution[];
};

// Initial, manually transcribed source: eight official U15 match reports.
// These are NOT fetched from the AKA Goals API. Add an import/persistence
// workflow before expecting newly entered matches to appear automatically.
const FIXTURES: Fixture[] = [
  {
    date: "2026-08-29", opponent: "Admira Wacker", result: "3:3",
    goals: [[19,"for"],[26,"against"],[49,"for"],[57,"for"],[66,"against"],[72,"against"]],
    starters: ["Tobias Sommer","Arnes Moric","David Motz","Simon Scharinger","Nail Kovacevic","Matteo Bendra","Diyan Behric","Bennet Berg","Bilal Tajroski","Moritz Breitauer","Maximilian Holz"],
    substitutions: [[61,"Tobias Sommer","Adnan Selmanovic"],[53,"Arnes Moric","Sebastian Itzinger"],[40,"Matteo Bendra","Marko Buljan"],[53,"Diyan Behric","Emir Aktas"],[53,"Bennet Berg","Theo Schatzl"],[40,"Bilal Tajroski","Niko Malekin"],[53,"Moritz Breitauer","Milos Gnjatic"]]
  },
  {
    date: "2026-09-04", opponent: "St. Pölten", result: "0:4",
    goals: [[36,"against"],[65,"against"],[68,"against"],[77,"against"]],
    starters: ["Adnan Selmanovic","Arnes Moric","David Motz","Simon Scharinger","Nail Kovacevic","Niko Malekin","Diyan Behric","Bennet Berg","Bilal Tajroski","Moritz Breitauer","Maximilian Holz"],
    substitutions: [[60,"Adnan Selmanovic","Tobias Sommer"],[54,"Arnes Moric","Sebastian Itzinger"],[40,"Niko Malekin","Amil Jusic"],[54,"Bennet Berg","Moamen Mezzi"],[40,"Bilal Tajroski","Matteo Bendra"],[54,"Moritz Breitauer","Paul Sickinger"],[54,"Maximilian Holz","Emir Aktas"]]
  },
  {
    date: "2026-09-12", opponent: "Vorarlberg", result: "1:4",
    goals: [[10,"for"],[29,"against"],[59,"against"],[73,"against"],[80,"against"]],
    starters: ["Adnan Selmanovic","Arnes Moric","Theo Schatzl","Felix Berger","Nail Kovacevic","Matteo Bendra","Marko Buljan","Bennet Berg","Maximilian Holz","Moritz Breitauer","Diyan Behric"],
    substitutions: [[57,"Theo Schatzl","Niko Malekin"],[57,"Felix Berger","Amil Jusic"],[57,"Marko Buljan","Bilal Tajroski"],[57,"Bennet Berg","Leon Mijatovic"],[40,"Moritz Breitauer","Sebastian Itzinger"],[48,"Diyan Behric","Emir Aktas"]]
  },
  {
    date: "2026-09-19", opponent: "Burgenland", result: "3:2",
    goals: [[2,"against"],[6,"against"],[38,"for"],[44,"for"],[80,"for"]],
    starters: ["Tobias Sommer","Arnes Moric","Milos Gnjatic","Nail Kovacevic","Matteo Bendra","Marko Buljan","Bennet Berg","Bilal Tajroski","Moritz Breitauer","Maximilian Holz","Simon Scharinger"],
    substitutions: [[77,"Bennet Berg","Amil Jusic"],[77,"Bilal Tajroski","Moamen Mezzi"],[77,"Moritz Breitauer","Leon Mijatovic"],[57,"Maximilian Holz","Emir Aktas"],[72,"Simon Scharinger","Niko Malekin"]]
  },
  {
    date: "2026-10-10", opponent: "Red Bull Salzburg", result: "1:9",
    goals: [[5,"against"],[28,"against"],[34,"against"],[38,"against"],[47,"against"],[49,"against"],[50,"for"],[61,"against"],[67,"against"],[72,"against"]],
    starters: ["Tobias Sommer","Marko Buljan","David Motz","Matteo Bendra","Anton Bauernfeind","Sebastian Itzinger","Sirac Celepci","Niko Malekin","Amil Jusic","Emir Aktas","Diyan Behric"],
    substitutions: [[60,"Tobias Sommer","Adnan Selmanovic"],[40,"David Motz","Moamen Mezzi"],[60,"Sirac Celepci","Nico Hartl"],[40,"Emir Aktas","Din Delibegovic"],[56,"Diyan Behric","Leon Mijatovic"]]
  },
  {
    date: "2026-10-17", opponent: "LASK", result: "1:5",
    goals: [[14,"against"],[21,"against"],[46,"for"],[67,"against"],[75,"against"],[80,"against"]],
    starters: ["Adnan Selmanovic","Anton Bauernfeind","David Motz","Simon Scharinger","Matteo Bendra","Sebastian Itzinger","Din Delibegovic","Bennet Berg","Nico Hartl","Moritz Breitauer","Diyan Behric"],
    substitutions: [[59,"Sebastian Itzinger","Moamen Mezzi"],[37,"Din Delibegovic","Emir Aktas"],[68,"Bennet Berg","Leon Mijatovic"],[68,"Nico Hartl","Amil Jusic"],[37,"Diyan Behric","Marko Buljan"]]
  },
  {
    date: "2026-10-24", opponent: "Austria Wien", result: "0:2",
    goals: [[34,"against"],[69,"against"]],
    starters: ["Tobias Sommer","Arnes Moric","Milos Gnjatic","Simon Scharinger","Anton Bauernfeind","Bennet Berg","Marko Buljan","David Motz","Nico Hartl","Moritz Breitauer","Emir Aktas"],
    substitutions: [[66,"Moritz Breitauer","Felix Berger"],[66,"Bennet Berg","Moamen Mezzi"],[66,"Simon Scharinger","Niko Malekin"],[53,"Marko Buljan","Diyan Behric"],[53,"Emir Aktas","Sirac Celepci"]]
  },
  {
    date: "2026-10-28", opponent: "AKA Steiermark / Sturm Graz", result: "1:6",
    goals: [[18,"against"],[25,"against"],[26,"against"],[45,"against"],[66,"against"],[70,"against"],[79,"for"]],
    starters: ["Tobias Sommer","Anton Bauernfeind","David Motz","Felix Berger","Niko Malekin","Bennet Berg","Diyan Behric","Moritz Breitauer","Amil Jusic","Moamen Mezzi","Emir Aktas"],
    substitutions: [[56,"Diyan Behric","Din Delibegovic"],[56,"Amil Jusic","Sirac Celepci"],[56,"Moamen Mezzi","Sebastian Itzinger"],[61,"Tobias Sommer","Adnan Selmanovic"]]
  }
];

const U16_ABSENCES: Record<string, string[]> = {
  "2026-09-12": ["Milos Gnjatic"],
  "2026-10-10": ["Nail Kovacevic","Milos Gnjatic","Arnes Moric","Simon Scharinger","Bennet Berg","Moritz Breitauer","Bilal Tajroski","Felix Berger"],
  "2026-10-17": ["Nail Kovacevic","Milos Gnjatic","Arnes Moric","Bilal Tajroski"],
  "2026-10-24": ["Nail Kovacevic"],
  "2026-10-28": ["Nail Kovacevic","Arnes Moric","Simon Scharinger","Milos Gnjatic","Marko Buljan","Bilal Tajroski"]
};


type PlayerRow = {
  name: string; minutes: number; appearances: number; starts: number;
  scored: number; conceded: number; plusMinus: number;
  jokerMinutes: number; jokerScored: number; jokerConceded: number;
  offMinutes: number; offScored: number; offConceded: number;
};
type ChangeRow = {
  date: string; opponent: string; minute: number; off: string; on: string;
  beforeFor: number; beforeAgainst: number;
  afterFor: number; afterAgainst: number;
  scoreAtChange: string;
};
type OffImpactRow = {
  date: string; opponent: string; minute: number; on: string;
  concurrent: number; beforeMinutes: number; afterMinutes: number;
  beforeFor: number; beforeAgainst: number; afterFor: number; afterAgainst: number;
};
type SubstitutionSummary = {
  name: string; count: number; beforeMinutes: number; afterMinutes: number;
  beforeFor: number; beforeAgainst: number; afterFor: number; afterAgainst: number;
  changePer80: number;
};

type U16AbsenceImpactRow = {
  name: string;
  u16Absences: number;
  withGames: number;
  withFor: number;
  withAgainst: number;
  withoutGames: number;
  withoutFor: number;
  withoutAgainst: number;
  withDiffPerGame: number;
  withoutDiffPerGame: number;
  availabilityImpact: number;
};


const signed = (n: number) => n > 0 ? `+${n}` : String(n);
const per80 = (n: number, mins: number) => mins ? (n * 80 / mins).toFixed(2).replace(".", ",") : "–";

function getOutgoingImpact(fixtures: Fixture[], name: string, range: "full" | "15"): OffImpactRow[] {
  return fixtures.flatMap(match => match.substitutions
    .filter(([, off]) => off === name)
    .map(([minute, , on]) => {
      const beforeStart = range === "full" ? 0 : Math.max(0, minute - 15);
      const afterEnd = range === "full" ? 80 : Math.min(80, minute + 15);
      const before = match.goals.filter(([m]) => m <= minute && m > beforeStart);
      const after = match.goals.filter(([m]) => m > minute && m <= afterEnd);
      return {
        date: match.date, opponent: match.opponent, minute, on,
        concurrent: match.substitutions.filter(([m]) => m === minute).length,
        beforeMinutes: minute-beforeStart, afterMinutes: afterEnd-minute,
        beforeFor: before.filter(([,s]) => s === "for").length,
        beforeAgainst: before.filter(([,s]) => s === "against").length,
        afterFor: after.filter(([,s]) => s === "for").length,
        afterAgainst: after.filter(([,s]) => s === "against").length
      };
    }));
}

function getSubstitutionSummaries(fixtures: Fixture[], direction: "incoming" | "outgoing"): SubstitutionSummary[] {
  const byName = new Map<string, Omit<SubstitutionSummary, "changePer80">>();
  for (const match of fixtures) {
    for (const [minute, off, on] of match.substitutions) {
      const name = direction === "incoming" ? on : off;
      const current = byName.get(name) ?? {
        name, count: 0, beforeMinutes: 0, afterMinutes: 0,
        beforeFor: 0, beforeAgainst: 0, afterFor: 0, afterAgainst: 0
      };
      current.count++;
      current.beforeMinutes += minute;
      current.afterMinutes += 80-minute;
      for (const [goalMinute, side] of match.goals) {
        // Goals in the substitution minute are attributed to the period before.
        if (goalMinute <= minute) {
          if (side === "for") current.beforeFor++;
          else current.beforeAgainst++;
        } else {
          if (side === "for") current.afterFor++;
          else current.afterAgainst++;
        }
      }
      byName.set(name, current);
    }
  }
  return [...byName.values()].map(p => ({
    ...p,
    changePer80: p.beforeMinutes > 0 && p.afterMinutes > 0
      ? ((p.afterFor-p.afterAgainst)/p.afterMinutes - (p.beforeFor-p.beforeAgainst)/p.beforeMinutes)*80
      : 0
  })).sort((a,b) => b.changePer80-a.changePer80 || b.count-a.count);
}

function getFixturePlayers(match: Fixture): Set<string> {
  return new Set([
    ...match.starters,
    ...match.substitutions.map(([, , on]) => on)
  ]);
}

function getU16AbsenceCount(fixtures: Fixture[], name: string): number {
  return fixtures.filter(match => (U16_ABSENCES[match.date] ?? []).includes(name)).length;
}

function getU16AbsenceImpact(fixtures: Fixture[]): U16AbsenceImpactRow[] {
  const names = [...new Set(
    fixtures.flatMap(match => U16_ABSENCES[match.date] ?? [])
  )];

  return names.map(name => {
    let u16Absences = 0;
    let withGames = 0;
    let withFor = 0;
    let withAgainst = 0;
    let withoutGames = 0;
    let withoutFor = 0;
    let withoutAgainst = 0;

    for (const match of fixtures) {
      const players = getFixturePlayers(match);
      const goalsFor = match.goals.filter(([, side]) => side === "for").length;
      const goalsAgainst = match.goals.filter(([, side]) => side === "against").length;
      const absentForU16 = (U16_ABSENCES[match.date] ?? []).includes(name);

      if (absentForU16) {
        u16Absences++;
        withoutGames++;
        withoutFor += goalsFor;
        withoutAgainst += goalsAgainst;
      } else if (players.has(name)) {
        withGames++;
        withFor += goalsFor;
        withAgainst += goalsAgainst;
      }
    }

    const withDiffPerGame = withGames > 0
      ? (withFor - withAgainst) / withGames
      : 0;

    const withoutDiffPerGame = withoutGames > 0
      ? (withoutFor - withoutAgainst) / withoutGames
      : 0;

    return {
      name,
      u16Absences,
      withGames,
      withFor,
      withAgainst,
      withoutGames,
      withoutFor,
      withoutAgainst,
      withDiffPerGame,
      withoutDiffPerGame,
      availabilityImpact: withDiffPerGame - withoutDiffPerGame
    };
  }).sort((a,b) =>
    b.availabilityImpact - a.availabilityImpact ||
    b.u16Absences - a.u16Absences ||
    a.name.localeCompare(b.name, "de")
  );
}

function computeImpact(fixtures: Fixture[]) {
  const players = new Map<string, PlayerRow>();
  const changes: ChangeRow[] = [];
  for (const match of fixtures) {
    const ranges = new Map<string, [number, number]>(
      match.starters.map(name => [name, [0, 80]])
    );
    for (const [minute, off, on] of match.substitutions) {
      const oldRange = ranges.get(off);
      if (!oldRange) throw new Error(`Unknown starter/substitution: ${off} (${match.date})`);
      oldRange[1] = minute;
      ranges.set(on, [minute, 80]);
      const before = match.goals.filter(([m]) => m < minute);
      const after = match.goals.filter(([m]) => m > minute);
      changes.push({
        date: match.date, opponent: match.opponent, minute, off, on,
        beforeFor: before.filter(([,s]) => s === "for").length,
        beforeAgainst: before.filter(([,s]) => s === "against").length,
        afterFor: after.filter(([,s]) => s === "for").length,
        afterAgainst: after.filter(([,s]) => s === "against").length,
        scoreAtChange: `${before.filter(([,s]) => s === "for").length}:${before.filter(([,s]) => s === "against").length}`
      });
    }
    for (const [name, [start, end]] of ranges) {
      const p = players.get(name) ?? {
        name, minutes: 0, appearances: 0, starts: 0,
        scored: 0, conceded: 0, plusMinus: 0,
        jokerMinutes: 0, jokerScored: 0, jokerConceded: 0,
        offMinutes: 0, offScored: 0, offConceded: 0
      };
      p.minutes += end - start;
      p.appearances++;
      if (start === 0) p.starts++;
      else p.jokerMinutes += end - start;
      for (const [minute, side] of match.goals) {
        // Match report minutes are treated as the end of the minute:
        // goals at the substitution minute belong to the outgoing player.
        const on = minute > start && minute <= end;
        if (on) {
          p.plusMinus += side === "for" ? 1 : -1;
          if (side === "for") p.scored++;
          else p.conceded++;
          if (start > 0) {
            if (side === "for") p.jokerScored++;
            else p.jokerConceded++;
          }
        }
      }
      players.set(name, p);
    }
    // On/off is only meaningful for players actually fielded in a match:
    // exclude absent matches, and compare their bench minutes within appearances.
    for (const [name, [start, end]] of ranges) {
      const p = players.get(name)!;
      p.offMinutes += start + (80 - end);
      for (const [minute, side] of match.goals) {
        if (minute <= start || minute > end) {
          if (side === "for") p.offScored++;
          else p.offConceded++;
        }
      }
    }
  }
  return { players: [...players.values()], changes };
}

export default function PlayerImpact() {
  const [selectedDates, setSelectedDates] = useState<string[]>(FIXTURES.map(f => f.date));
  const [tab, setTab] = useState<"players" | "jokers" | "changes" | "onoff" | "outgoing" | "substitution-summary" | "u16-absence">("players");
  const [minMinutes, setMinMinutes] = useState(80);
  const [selectedPlayer, setSelectedPlayer] = useState("");
  const [outgoingRange, setOutgoingRange] = useState<"full" | "15">("full");
  const [minSubstitutions, setMinSubstitutions] = useState(1);
  const fixtures = useMemo(() => FIXTURES.filter(f => selectedDates.includes(f.date)), [selectedDates]);
  const { players, changes } = useMemo(() => computeImpact(fixtures), [fixtures]);
  const ranked = [...players]
    .filter(p => {
      const minutes = tab === "jokers" ? p.jokerMinutes : p.minutes;
      return minutes > 0 && minutes >= minMinutes;
    })
    .sort((a,b) => {
      const am = tab === "jokers" ? a.jokerMinutes : a.minutes;
      const bm = tab === "jokers" ? b.jokerMinutes : b.minutes;
      const av = tab === "jokers" ? a.jokerScored - a.jokerConceded : a.plusMinus;
      const bv = tab === "jokers" ? b.jokerScored - b.jokerConceded : b.plusMinus;
      return (bv / bm) - (av / am);
    });
  const detail = players.find(p => p.name === selectedPlayer);
  const outgoingNames = [...new Set(fixtures.flatMap(f => f.substitutions.map(([,off]) => off)))].sort();
  const outgoingName = outgoingNames.includes(selectedPlayer) ? selectedPlayer : (outgoingNames.includes("Bennet Berg") ? "Bennet Berg" : outgoingNames[0]);
  const outgoingRows = outgoingName ? getOutgoingImpact(fixtures, outgoingName, outgoingRange) : [];
  const outgoingTotals = outgoingRows.reduce((s,r) => ({
    beforeMinutes:s.beforeMinutes+r.beforeMinutes, afterMinutes:s.afterMinutes+r.afterMinutes,
    beforeFor:s.beforeFor+r.beforeFor, beforeAgainst:s.beforeAgainst+r.beforeAgainst,
    afterFor:s.afterFor+r.afterFor, afterAgainst:s.afterAgainst+r.afterAgainst
  }), {beforeMinutes:0,afterMinutes:0,beforeFor:0,beforeAgainst:0,afterFor:0,afterAgainst:0});
  const incomingSummary = useMemo(() => getSubstitutionSummaries(fixtures, "incoming"), [fixtures]);
  const outgoingSummary = useMemo(() => getSubstitutionSummaries(fixtures, "outgoing"), [fixtures]);
  const u16AbsenceImpact = useMemo(() => getU16AbsenceImpact(fixtures), [fixtures]);
  const totalFor = fixtures.reduce((n,f) => n + f.goals.filter(g => g[1] === "for").length, 0);
  const totalAgainst = fixtures.reduce((n,f) => n + f.goals.filter(g => g[1] === "against").length, 0);
  const cell = {padding:"10px", borderBottom:"1px solid #e1e5e2", textAlign:"left" as const};

  return (
    <section style={{maxWidth:1200,margin:"0 auto"}}>
      <h2>U15 – Spieler-Impact</h2>
      <p style={{color:"#526052"}}>
        Ergebnis-Impact aus offiziellen Spielberichten (Stand 28.10.2026).
        Kein IMPECT-Tracking und keine automatische Übernahme neuer AKA-Goals-Ereignisse.
      </p>
      <div className="analysis-kpis">
        <div className="kpi-card"><span>Spiele</span><strong>{fixtures.length}</strong></div>
        <div className="kpi-card"><span>Tore</span><strong>{totalFor}</strong></div>
        <div className="kpi-card"><span>Gegentore</span><strong>{totalAgainst}</strong></div>
        <div className="kpi-card"><span>Bilanz</span><strong>{signed(totalFor-totalAgainst)}</strong></div>
      </div>
      <div style={{background:"#fff",border:"1px solid #ddd",borderRadius:12,padding:16,margin:"20px 0"}}>
        <strong>Spiele berücksichtigen</strong>
        <div style={{display:"flex",flexWrap:"wrap",gap:12,marginTop:12}}>
          {FIXTURES.map(f => <label key={f.date} style={{display:"flex",gap:5,alignItems:"center",fontSize:14}}>
            <input type="checkbox" checked={selectedDates.includes(f.date)}
              onChange={e => setSelectedDates(current => e.target.checked ? [...current,f.date] : current.filter(d => d !== f.date))}/>
            {f.date.slice(5)} · {f.opponent} ({f.result})
          </label>)}
        </div>
      </div>
      <div className="view-switcher" style={{flexWrap:"wrap",marginBottom:18}}>
        {([["players","Spieler-Ranking"],["jokers","Joker-Impact"],["changes","Wechsel-Impact"],["onoff","On/Off"],["outgoing","Nach Auswechslung"],["substitution-summary","Wechsel-Gesamtübersicht"],["u16-absence","U16-Abwesenheits-Impact"]] as const)
          .map(([key,label]) => <button key={key} className={tab===key?"view-button-active":""} onClick={() => setTab(key)}>{label}</button>)}
      </div>
      {tab !== "changes" && tab !== "outgoing" && tab !== "substitution-summary" && tab !== "u16-absence" && <div style={{display:"flex",gap:12,alignItems:"center",marginBottom:12,flexWrap:"wrap"}}>
        <label>Mindestminuten: <select value={minMinutes} onChange={e=>setMinMinutes(Number(e.target.value))}>
          {[0,40,80,160,240].map(m=><option key={m} value={m}>{m}</option>)}
        </select></label>
        <span style={{fontSize:13,color:"#526052"}}>Sortierung: Plus/Minus je 80 Einsatzminuten</span>
      </div>}
      {tab === "u16-absence" ? (
        <div style={{display:"grid",gap:18}}>
          <div style={{background:"#fff",borderRadius:12,padding:16}}>
            <h3>U16-Abwesenheits-Impact</h3>
            <p style={{fontSize:13,color:"#526052",marginBottom:0}}>
              Verglichen wird die U15-Gesamtbilanz in Spielen, in denen der Spieler eingesetzt wurde,
              mit den Spielen, die er nachweislich wegen eines U16-Einsatzes verpasst hat.
              Ein positiver Wert bedeutet: Die U15-Tordifferenz pro Spiel war mit dem Spieler besser.
            </p>
          </div>

          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",background:"#fff"}}>
              <thead>
                <tr>
                  {["Spieler","U16-Abstellungen","U15 mit Spieler","Bilanz mit","Diff./Spiel mit","U15 ohne wegen U16","Bilanz ohne","Diff./Spiel ohne","Availability Impact"].map(h=><th key={h} style={cell}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {u16AbsenceImpact.map(r=><tr key={r.name}>
                  <td style={cell}>{r.name}</td>
                  <td style={cell}>{r.u16Absences}</td>
                  <td style={cell}>{r.withGames}</td>
                  <td style={cell}>{r.withFor}:{r.withAgainst}</td>
                  <td style={cell}>{signed(Number(r.withDiffPerGame.toFixed(2))).replace(".",",")}</td>
                  <td style={cell}>{r.withoutGames}</td>
                  <td style={cell}>{r.withoutFor}:{r.withoutAgainst}</td>
                  <td style={cell}>{signed(Number(r.withoutDiffPerGame.toFixed(2))).replace(".",",")}</td>
                  <td style={{...cell,fontWeight:700,color:r.availabilityImpact>0?"#187a38":r.availabilityImpact<0?"#b32b2b":undefined}}>
                    {signed(Number(r.availabilityImpact.toFixed(2))).replace(".",",")}
                  </td>
                </tr>)}
              </tbody>
            </table>
            {!u16AbsenceImpact.length&&<p>Für die gewählte Spielauswahl sind keine bestätigten U16-Abstellungen hinterlegt.</p>}
          </div>

          <p style={{fontSize:13,color:"#526052"}}>
            Wichtig: Das ist kein kausaler Einzelspielerwert. Besonders Spiele wie Red Bull, in denen mehrere U15-Spieler gleichzeitig in der U16 waren,
            beeinflussen mehrere Spieler mit demselben Mannschaftsergebnis. Der Wert zeigt deshalb einen Availability-/Abwesenheitszusammenhang.
          </p>
        </div>
      ) : tab === "substitution-summary" ? (
        <div style={{display:"grid",gap:22}}>
          <div style={{background:"#fff",borderRadius:12,padding:16}}>
            <label>Mindestens <select value={minSubstitutions} onChange={e=>setMinSubstitutions(Number(e.target.value))}>
              {[1,2,3,4].map(n=><option key={n} value={n}>{n} Wechsel</option>)}
            </select> pro Spieler</label>
            <p style={{fontSize:13,color:"#526052"}}>Für jeden Wechsel werden der Abschnitt vom Anpfiff bis zum Wechsel und der Abschnitt vom Wechsel bis zur 80. Minute verglichen. Die Rangfolge basiert auf der Veränderung der Tordifferenz pro 80 Minuten.</p>
          </div>
          {([[
            "Einwechsel-Impact", incomingSummary,
            "Ergebnisentwicklung der Mannschaft vor und nach der Einwechslung"
          ],[
            "Auswechsel-Impact", outgoingSummary,
            "Ergebnisentwicklung der Mannschaft vor und nach der Auswechslung"
          ]] as const).map(([title, rows, description]) => (
            <div key={title} style={{background:"#fff",borderRadius:12,padding:16}}>
              <h3>{title}</h3>
              <p style={{fontSize:13,color:"#526052"}}>{description}</p>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr>{["Spieler","Wechsel","Vorher Minuten","Vorher Tore","Vorher +/-","Nachher Minuten","Nachher Tore","Nachher +/-","Veränderung / 80 Min."].map(h=><th key={h} style={cell}>{h}</th>)}</tr></thead>
                  <tbody>{rows.filter(r=>r.count>=minSubstitutions).map(r=><tr key={r.name}>
                    <td style={cell}>{r.name}</td><td style={cell}>{r.count}</td>
                    <td style={cell}>{r.beforeMinutes}</td><td style={cell}>{r.beforeFor}:{r.beforeAgainst}</td>
                    <td style={cell}>{signed(r.beforeFor-r.beforeAgainst)}</td>
                    <td style={cell}>{r.afterMinutes}</td><td style={cell}>{r.afterFor}:{r.afterAgainst}</td>
                    <td style={cell}>{signed(r.afterFor-r.afterAgainst)}</td>
                    <td style={{...cell,fontWeight:700,color:r.changePer80>0?"#187a38":r.changePer80<0?"#b32b2b":undefined}}>{signed(Number(r.changePer80.toFixed(2))).replace(".",",")}</td>
                  </tr>)}</tbody>
                </table>
                {!rows.some(r=>r.count>=minSubstitutions)&&<p>Keine Spieler mit ausreichend Wechseln.</p>}
              </div>
            </div>
          ))}
          <p style={{fontSize:13,color:"#526052"}}>Positive Veränderung bedeutet bessere Ergebnisentwicklung nach dem Wechsel, nicht zwingend positiven individuellen Einfluss. Gleichzeitige Wechsel erscheinen bei mehreren Spielern mit denselben Mannschaftswerten. Aufeinanderfolgende Wechsel erzeugen überlappende Zeiträume.</p>
        </div>
      ) : tab === "outgoing" ? (
        <div style={{background:"#fff",borderRadius:12,padding:16}}>
          <label style={{display:"flex",gap:12,alignItems:"center",flexWrap:"wrap",marginBottom:16}}>
            <strong>Ausgewechselter Spieler:</strong>
            <select value={outgoingName ?? ""} onChange={e=>setSelectedPlayer(e.target.value)}>
              {outgoingNames.map(n=><option key={n} value={n}>{n}</option>)}
            </select>
          </label>
          <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",marginBottom:12}}>
            <strong>Betrachtungszeitraum:</strong>
            <button className={outgoingRange === "full" ? "view-button-active" : ""} onClick={() => setOutgoingRange("full")}>Gesamtes Spiel</button>
            <button className={outgoingRange === "15" ? "view-button-active" : ""} onClick={() => setOutgoingRange("15")}>15 Minuten davor/danach</button>
          </div>
          <p>{outgoingRange === "full"
            ? "Vergleich vom Anpfiff bis zur Auswechslung und von der Auswechslung bis zum Spielende (80. Minute)."
            : "Vergleich der jeweils maximal 15 Minuten unmittelbar vor und nach der Auswechslung."}</p>
          <div className="analysis-kpis">
            <div className="kpi-card"><span>Auswechslungen</span><strong>{outgoingRows.length}</strong></div>
            <div className="kpi-card"><span>Vorher +/-</span><strong>{signed(outgoingTotals.beforeFor-outgoingTotals.beforeAgainst)}</strong></div>
            <div className="kpi-card"><span>Nachher +/-</span><strong>{signed(outgoingTotals.afterFor-outgoingTotals.afterAgainst)}</strong></div>
            <div className="kpi-card"><span>Veränderung / 80 Min.</span><strong>{outgoingTotals.afterMinutes && outgoingTotals.beforeMinutes ? signed(Number(((outgoingTotals.afterFor-outgoingTotals.afterAgainst)*80/outgoingTotals.afterMinutes - (outgoingTotals.beforeFor-outgoingTotals.beforeAgainst)*80/outgoingTotals.beforeMinutes).toFixed(2))) : "–"}</strong></div>
          </div>
          <div style={{overflowX:"auto",marginTop:16}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>{["Spiel","Minute","Eingewechselt","Vorher (Tore)","Nachher (Tore)","Vorher +/-","Nachher +/-","Weitere Wechsel"].map(h=><th key={h} style={cell}>{h}</th>)}</tr></thead>
              <tbody>{outgoingRows.map((r,i)=><tr key={i}>
                <td style={cell}>{r.date.slice(5)} · {r.opponent}</td><td style={cell}>{r.minute}′</td>
                <td style={cell}>{r.on}</td><td style={cell}>{r.beforeFor}:{r.beforeAgainst} ({r.beforeMinutes}′)</td>
                <td style={cell}>{r.afterFor}:{r.afterAgainst} ({r.afterMinutes}′)</td>
                <td style={cell}>{signed(r.beforeFor-r.beforeAgainst)}</td><td style={cell}>{signed(r.afterFor-r.afterAgainst)}</td>
                <td style={cell}>{r.concurrent>1 ? `${r.concurrent-1} gleichzeitig` : "keine gleichzeitig"}</td>
              </tr>)}</tbody>
            </table>
            {!outgoingRows.length&&<p>Für diese Spielauswahl liegen keine Auswechslungen vor.</p>}
          </div>
          <p style={{fontSize:13,color:"#526052"}}>Die Vorher/Nachher-Zeiträume sind unterschiedlich lang. Die Veränderung pro 80 Minuten gleicht die Dauer rechnerisch an, bleibt bei kurzen Restspielzeiten aber besonders unsicher. Andere Wechsel, Spielstand und Gegner beeinflussen die Werte. Eine kausale Wirkung des ausgewechselten Spielers lässt sich daraus nicht ableiten.</p>
        </div>
      ) : tab === "changes" ? (
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",background:"#fff"}}>
            <thead><tr>{["Spiel","Min.","Aus","Ein","Spielstand","Danach Tore","Danach +/-"].map(h=><th key={h} style={cell}>{h}</th>)}</tr></thead>
            <tbody>{changes.sort((a,b)=>a.date.localeCompare(b.date)||a.minute-b.minute).map((c,i)=>
              <tr key={i}><td style={cell}>{c.date.slice(5)} · {c.opponent}</td><td style={cell}>{c.minute}′</td>
              <td style={cell}>{c.off}</td><td style={cell}>{c.on}</td><td style={cell}>{c.scoreAtChange}</td>
              <td style={cell}>{c.afterFor}:{c.afterAgainst}</td><td style={cell}>{signed(c.afterFor-c.afterAgainst)}</td></tr>)}</tbody>
          </table>
          <p style={{fontSize:13,color:"#526052"}}>Bei zeitgleichen Wechseln ist der nachfolgende Teameffekt derselbe; er kann keinem einzelnen Joker kausal zugeschrieben werden.</p>
        </div>
      ) : (
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",background:"#fff"}}>
            <thead><tr>{["Spieler","Einsätze","U16-Abstellungen","Minuten","Tore mit","Gegentore mit","+/-","Pro 80 Min."].map(h=><th key={h} style={cell}>{h}</th>)}</tr></thead>
            <tbody>{ranked.map(p=>{
              const joker=tab==="jokers";
              const mins=joker?p.jokerMinutes:p.minutes;
              const gf=joker?p.jokerScored:p.scored;
              const ga=joker?p.jokerConceded:p.conceded;
              return <tr key={p.name} onClick={()=>setSelectedPlayer(p.name)} style={{cursor:"pointer",background:selectedPlayer===p.name?"#e9f7e8":undefined}}>
                <td style={cell}>{p.name}</td><td style={cell}>{p.appearances} ({p.starts} Startelf)</td>
                <td style={cell}>{getU16AbsenceCount(fixtures,p.name)}</td>
                <td style={cell}>{mins}</td><td style={cell}>{gf}</td><td style={cell}>{ga}</td>
                <td style={cell}>{signed(gf-ga)}</td><td style={cell}>{per80(gf-ga,mins)}</td></tr>;
            })}</tbody>
          </table>
          {ranked.length===0&&<p>Für diesen Filter liegen keine Spieler vor.</p>}
        </div>
      )}
      {tab==="onoff"&&<p style={{color:"#526052",fontSize:13}}>
        On/Off wird ausschließlich innerhalb der Spiele verglichen, in denen der Spieler tatsächlich eingesetzt wurde.
        Die ausgewiesenen Spielphasen sind durch Gegner, Spielstand, Mitspieler und Wechsel beeinflusst.
      </p>}
      {detail&&tab!=="changes"&&tab!=="outgoing"&&tab!=="substitution-summary"&&tab!=="u16-absence"&&<div style={{background:"#eef8ed",padding:16,borderRadius:12,marginTop:18}}>
        <strong>{detail.name} · On/Off innerhalb seiner {detail.appearances} Spiele</strong>
        <p>Auf dem Platz: {detail.minutes} Minuten, {detail.scored}:{detail.conceded}, +/- {signed(detail.plusMinus)}
          {" "}({per80(detail.plusMinus,detail.minutes)} je 80 Min.)</p>
        <p>Auf der Bank: {detail.offMinutes} Minuten, {detail.offScored}:{detail.offConceded},
          {" "}+/- {signed(detail.offScored-detail.offConceded)}
          {" "}({per80(detail.offScored-detail.offConceded,detail.offMinutes)} je 80 Min.)</p>
      </div>}
      <p style={{fontSize:13,color:"#526052",marginTop:20}}>
        Methodik: 80 Minuten pro Partie, Halbzeitwechsel bei Minute 40.
        Gleichzeitige Wechsel werden gemeinsam betrachtet. Bestätigte U16-Abstellungen werden separat
        als Abwesenheits-/Availability-Impact ausgewertet. Die Werte beschreiben Ergebnisphasen und Zusammenhänge,
        nicht die isolierte individuelle Leistung.
      </p>
    </section>
  );
}
