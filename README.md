# PPT Animation Transfer

A high-fidelity tool to automatically transfer **ALL animations and transitions** from an **ORIGINAL .pptx** to a **REMEDIATED .pptx**, without modifying any original presentation files on disk.

---

## Key Features

1. **Intelligent Slide Matching**:
   - Matches slides by index first.
   - Falls back to rapidfuzz token similarity for title + text when slide counts or order diverge.
   - Reports unmatched slides clearly.

2. **4-Stage Shape Mapping Cascade**:
   - Pass 1: Same Shape ID + same shape type + similar text.
   - Pass 2: Same Shape Name + same shape type.
   - Pass 3: Shape type + normalized text similarity (RapidFuzz) + spatial proximity score (bounding box distance).
   - Pass 4: Image MD5 hash for picture shapes.
   - Enforces strict 1-to-1 matching constraint (each remediated shape is matched only once).

3. **Complete XML Animation Tree Transfer**:
   - Preserves `<p:timing>`: main sequences, interactive sequences, click triggers, delays, durations, repeat counts, easing, motion paths (`<p:animMotion>`), rotations, scaling, and color animations.
   - Preserves by-paragraph text builds (`<p:bldP>`, `<p:txEl>`, `<p:pRg>`).
   - Deep rewrites of target shape IDs in `<p:spTgt spid="...">`, build list entries, and interactive sequence conditions.
   - Prunes orphaned animation nodes and cleans empty click containers.
   - Re-indexes `<p:cTn id="...">` sequentially to guarantee uniqueness.

4. **Optional Slide Transitions**:
   - Copies `<p:transition>` and `mc:AlternateContent` wrappers (e.g. PowerPoint 2010+ transitions) into remediated slides.

5. **Audio/Video Media Migration**:
   - Migrates embedded media relationships (`.rels`) for media-triggered animations.

6. **Validation & Verification Engine**:
   - Validates output integrity with `python-pptx`, XML schema/well-formedness, `cTn` uniqueness, and target shape existence.
   - Automatic shape-by-shape animation verification (Pass/Fail) comparing effect type, name, order, trigger, delay, and duration.

7. **Reports & Downloads**:
   - Download `<remediated_name>_animated.pptx`.
   - Download Excel report (.xlsx) with 3 sheets: `Summary`, `Animations`, `Skipped`.
   - Cached in `st.session_state` to prevent re-running on download.

---

## Quick Start

### 1. Launch the Web Interface
Double-click `run_app.bat` or run:
```bash
streamlit run app.py
```

### 2. Run Tests
```bash
pytest tests -v
```

---

## Final Manual Check in Microsoft PowerPoint
After downloading your generated presentation:
1. Open the file in **Microsoft PowerPoint**.
2. Press **F5** (or **Shift + F5**) to start the slideshow.
3. Advance through the slides and verify entrance, emphasis, exit, motion paths, and paragraph builds play in correct sequence.
