import json

with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\data.json", "r", encoding="utf-8") as f:
    db = json.load(f)

students = db["students"]

sections_head_map = {}
for s in students:
    sec = s["class_name"]
    hname = s["sec_head_name"]
    htitle = s["sec_head_title"]
    hsig = s["sec_head_sig"]
    sections_head_map[sec] = (hname, htitle, hsig)

print(f"Total sections: {len(sections_head_map)}")
print("=" * 80)
for idx, (sec, (hname, htitle, hsig)) in enumerate(sorted(sections_head_map.items())):
    print(f"{idx+1:02d}. {sec:32s} -> {hname:20s} | {htitle}")
