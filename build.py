import openpyxl, json, sys
# Standings (rank) and next round fixtures, transcribed from the screenshots (home team listed first)
TEAMS = {"PAO":("Panathinaikos",1),"VBC":("Valencia Basket",2),"FBT":("Fenerbahçe",3),"OLY":("Olympiacos",4),"CZV":("Crvena zvezda",5),
 "HTA":("Hapoel Tel Aviv",6),"BAR":("Barcelona",7),"PAR":("Partizan",8),"VIR":("Virtus Bologna",9),"DUB":("Dubai",10),
 "RMB":("Real Madrid",11),"BJK":("Beşiktaş",12),"ZAL":("Žalgiris",13),"BAY":("Bayern Munich",14),"ASV":("LDLC ASVEL",15),
 "PBB":("Paris",16),"EFS":("Anadolu Efes",17),"MTA":("Maccabi Tel Aviv",18),"MIL":("Milano",19),"KBA":("Baskonia",20)}
# Next round (6) fixtures, home first; first 4 games are Turn 1 (15/10), last 6 are Turn 2 (16/10)
FIXTURES = [("FBT","PAR"),("MIL","DUB"),("VBC","MTA"),("RMB","ASV"),
            ("BJK","BAY"),("ZAL","EFS"),("CZV","PBB"),("PAO","HTA"),("BAR","OLY"),("VIR","KBA")]
# Rounds 1-5 fixtures (home first), transcribed from screenshots. Round 4 last game (BAR-ZAL) was cut off in the
# screenshot and inferred as the two teams not otherwise listed; its home/away order is unknown.
ROUNDS = {  # round number -> fixtures; add a new entry each round
 1: [("DUB","RMB"),("HTA","BAY"),("CZV","ZAL"),("PAO","PBB"),("BAR","EFS"),("KBA","OLY"),("ASV","MTA"),("BJK","VBC"),("FBT","VIR"),("PAR","MIL")],
 2: [("DUB","BAR"),("EFS","RMB"),("ZAL","OLY"),("FBT","BAY"),("CZV","HTA"),("VBC","KBA"),("MIL","VIR"),("PBB","PAR"),("MTA","BJK"),("PAO","ASV")],
 3: [("HTA","RMB"),("CZV","EFS"),("VIR","OLY"),("PBB","ZAL"),("BJK","BAR"),("FBT","DUB"),("ASV","VBC"),("BAY","PAR"),("PAO","MTA"),("KBA","MIL")],
 4: [("PBB","ASV"),("MTA","MIL"),("DUB","CZV"),("BAY","VIR"),("PAO","FBT"),("VBC","HTA"),("RMB","PAR"),("OLY","EFS"),("KBA","BJK"),("BAR","ZAL",None)],
 5: [("FBT","ZAL"),("ASV","CZV"),("KBA","DUB"),("MIL","RMB"),("BAR","MTA"),("PAR","PAO"),("VBC","OLY"),("HTA","BJK"),("EFS","BAY"),("PBB","VIR")],
}
OPP = {}
for rn_, fx in ROUNDS.items():
    m = {}
    for h, a, *unk in fx:
        known = not unk  # a trailing None marks an unknown home/away order
        m[h] = (a, "H" if known else ""); m[a] = (h, "A" if known else "")
    assert len(m) == 20, "every team must appear exactly once per round"
    OPP[rn_] = m
FIXTURE_TURNS = [1]*4 + [2]*6  # per FIXTURES entry: 4 games in Turn 1, 6 in Turn 2
NEXT = {}; TURN = {}
for (h, a), t in zip(FIXTURES, FIXTURE_TURNS):
    NEXT[h] = (a, "H"); NEXT[a] = (h, "A"); TURN[h] = TURN[a] = t
src = sys.argv[1]
import glob
def key(r): return r[0] if isinstance(r[0], int) else f"{r[1]} {r[2]}|{r[5]}"
rounds = {}
for path in sys.argv[2:]:
    w = openpyxl.load_workbook(path, data_only=True).active
    rn_ = next(int(str(r[0]).split(":")[1]) for r in w.iter_rows(values_only=True) if str(r[0]).startswith("Reference matchdays"))
    rounds[rn_] = ({key(r): (float(r[6]), float(r[7])) for r in w.iter_rows(min_row=2, values_only=True) if r[4] is not None})
# hover charts show only the last 5 rounds
LAST = sorted(rounds)[-5:]
ws = openpyxl.load_workbook(src, data_only=True).active
ref = ""
players = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[4] is None:
        if r[0]: ref = str(r[0])
        continue
    fpt, price, plus = float(r[6]), float(r[7]), float(r[8])
    players.append(dict(name=f"{r[1]} {r[2]}".strip(), pos=r[4], team=r[5], fpt=fpt, price=price, plus=plus,
                        fptpm=round(fpt/price, 3) if price else 0, pluspm=round(plus/price, 3) if price else 0,
                        rn=LAST, opps=[OPP[n].get(r[5]) for n in LAST], nextopp=TEAMS[NEXT[r[5]][0]][0], nextrank=TEAMS[NEXT[r[5]][0]][1], nextha=NEXT[r[5]][1], nextturn="T%d" % TURN[r[5]],
                        hist=[list(rounds[n][key(r)]) if key(r) in rounds[n] else None for n in LAST]))
import base64
logo = "data:image/png;base64," + base64.b64encode(open("euroleague-logo.png","rb").read()).decode()
html = open("template.html").read().replace("__LOGO__", logo).replace("__DATA__", json.dumps(players)).replace("__REF__", ref)
open("fantasy.html", "w").write(html)
print(len(players), "players")
