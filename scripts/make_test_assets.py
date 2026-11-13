import json, os, subprocess, zipfile
from PIL import Image, ImageDraw, ImageFont
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "tests", "assets"); os.makedirs(OUT, exist_ok=True)
DIM = {"1x1":(1200,1200),"4x5":(1080,1350),"9x16":(1080,1920),"191x1":(1200,628),"16x9":(1920,1080),"4x1":(1200,300)}
COL = {"FOUNDER_STORY":"#2C55E0","AGENCY_OWNER":"#1E9E6A","BENEFITS_STATIC":"#7A4FD6","DEMO_WALKTHROUGH":"#C2410C",
       "CASE_STUDY":"#0E7490","PROBLEM_SOLUTION":"#B91C1C","AGENCY_RESULTS":"#4D7C0F","TESTIMONIAL":"#9D174D","BRAND_LOGO":"#334155","BEFORE_AFTER_WORKFLOW":"#A16207"}
B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"; R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
def frame(name, con, ratio, kind):
    w, h = DIM[ratio]; im = Image.new("RGB", (w, h), COL[con]); d = ImageDraw.Draw(im); u = min(w, h)
    band = int(u*0.16); y0 = h//2 - band//2
    d.rectangle([0, y0, w, y0+band], fill="#FFD400")
    f1 = ImageFont.truetype(B, int(band*0.42)); t = "TEST FILE – NOT APPROVED"
    while d.textlength(t, font=f1) > w*0.92: f1 = ImageFont.truetype(B, f1.size-2)
    d.text((w/2, y0+band/2), t, font=f1, fill="#111111", anchor="mm")
    f2 = ImageFont.truetype(B, max(14, int(u*0.06))); f3 = ImageFont.truetype(R, max(12, int(u*0.035)))
    d.text((w/2, y0 - u*0.08), con.replace("_"," "), font=f2, fill="white", anchor="mm")
    d.text((w/2, y0 + band + u*0.07), f"{kind} · {ratio.replace('191x1','1.91:1').replace('x',':')} · {w}×{h}px", font=f3, fill="white", anchor="mm")
    d.text((w/2, h - u*0.05), name, font=f3, fill="white", anchor="mm")
    return im
for name, con, ver, ratio, kind in json.load(open(os.path.join(ROOT, "build", "library_spec.json"))):
    im = frame(name, con, ratio, kind); path = os.path.join(OUT, name)
    if kind == "Video":
        png = path + ".frame.png"; im.save(png)
        subprocess.run(["ffmpeg","-y","-loglevel","error","-loop","1","-i",png,"-t","12","-r","24","-c:v","libx264","-preset","veryfast",
                        "-crf","35","-pix_fmt","yuv420p",path], check=True); os.remove(png)
    else: im.save(path, optimize=True)
open(os.path.join(OUT,"README_TEST_ASSETS.txt"),"w").write(
 "These files are GENERATED STAND-INS for testing the campaign generator's file checker and mapping.\n"
 "Every file is stamped 'TEST FILE – NOT APPROVED' and its name ends in _TEST.\nDo not upload them to a live ad account.\n")
print(len(os.listdir(OUT)), "test files written to tests/assets")
