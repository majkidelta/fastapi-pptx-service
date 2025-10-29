from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from io import BytesIO

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"],
)

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/template/analyze")
async def analyze_template(file: UploadFile = File(...)):
    data = await file.read()
    prs = Presentation(BytesIO(data))

    layouts = []
    for i, layout in enumerate(prs.slide_layouts):
        name = getattr(layout, "name", None) or f"Layout {i}"
        placeholders = []
        for ph in layout.placeholders:
            ph_idx = getattr(getattr(ph, "placeholder_format", None), "idx", None)
            placeholders.append({
                "index": ph_idx,
                "type": "other",
                "name": getattr(ph, "name", None)
            })
        layouts.append({"name": name, "index": i, "placeholders": placeholders})

    EMU_PER_INCH = 914400
    DPI = 96
    slide_w_px = int((prs.slide_width / EMU_PER_INCH) * DPI)
    slide_h_px = int((prs.slide_height / EMU_PER_INCH) * DPI)

    style_profile = {
        "slideWidthPx": slide_w_px,
        "slideHeightPx": slide_h_px,
        "themeColors": [],
        "fonts": {"fallback": ["Arial", "Calibri"]},
        "layouts": layouts
    }
    return {"styleProfile": style_profile}
