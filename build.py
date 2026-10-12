import openpyxl, json, sys
# Standings (rank) and next round fixtures, transcribed from the screenshots (home team listed first)
TEAMS = {"PAO":("Panathinaikos",1),"VBC":("Valencia Basket",2),"OLY":("Olympiacos",3),"FBT":("Fenerbahçe",4),"PAR":("Partizan",5),
 "CZV":("Crvena zvezda",6),"BJK":("Beşiktaş",7),"ZAL":("Žalgiris",8),"BAR":("Barcelona",9),"VIR":("Virtus Bologna",10),
 "HTA":("Hapoel Tel Aviv",11),"DUB":("Dubai",12),"ASV":("LDLC ASVEL",13),"RMB":("Real Madrid",14),"EFS":("Anadolu Efes",15),
 "PBB":("Paris",16),"BAY":("Bayern Munich",17),"MTA":("Maccabi Tel Aviv",18),"KBA":("Baskonia",19),"MIL":("Milano",20)}
FIXTURES = [("FBT","ZAL"),("ASV","CZV"),("KBA","DUB"),("MIL","RMB"),("BAR","MTA"),("PAR","PAO"),("VBC","OLY"),
            ("HTA","BJK"),("EFS","BAY"),("PBB","VIR")]
NEXT = {}
for h, a in FIXTURES:
    NEXT[h] = (a, "H"); NEXT[a] = (h, "A")
src = sys.argv[1]
import glob
def key(r): return r[0] if isinstance(r[0], int) else f"{r[1]} {r[2]}|{r[5]}"
rounds = []
for path in sys.argv[2:]:
    w = openpyxl.load_workbook(path, data_only=True).active
    rounds.append({key(r): (float(r[6]), float(r[7])) for r in w.iter_rows(min_row=2, values_only=True) if r[4] is not None})
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
                        nextopp=TEAMS[NEXT[r[5]][0]][0], nextrank=TEAMS[NEXT[r[5]][0]][1], nextha=NEXT[r[5]][1],
                        hist=[list(rd[key(r)]) if key(r) in rd else None for rd in rounds]))
import base64
logo = "data:image/png;base64," + base64.b64encode(open("euroleague-logo.png","rb").read()).decode()
html = open("template.html").read().replace("__LOGO__", logo).replace("__DATA__", json.dumps(players)).replace("__REF__", ref)
open("fantasy.html", "w").write(html)
print(len(players), "players")
