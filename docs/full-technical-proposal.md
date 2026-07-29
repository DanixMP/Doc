بسمه تعالی

# Technology Development Proposal  
## Intelligent Cytology Screening System for Cervical Cancer (Pap Smear + AI)

**Form type:** فرم درخواست طرح توسعه فناوری  
**Duration:** 12 months  
**Domains:** Medical devices / Health + ICT  
**Outputs:** Product prototype + Native dataset  

---

# Part A — Official Form Structure

## 1. Project Title
Development of an intelligent system for identifying and classifying cytological abnormalities in Pap smear samples for cervical cancer diagnosis in women using artificial intelligence.

## 2. Exact Project Summary (≤ 300 words)
Cervical cancer is one of the most preventable cancers in women. Early detection reduces mortality and treatment cost. Pap smear screening still depends on specialist pathologists and suffers from human error, fatigue, specialist shortage, and inconsistent interpretation—especially at national scale.

This project builds a **native, on-premise clinical decision-support system** that:
1. Reads **whole-slide images (WSI)** of digital Pap smears directly (no manual cell cropping).
2. Detects and segments cells / nuclei.
3. Classifies nuclei as **normal vs abnormal**.
4. Fuses **clinical data** (age, prior treatment, symptoms, risk factors) into the final decision.
5. Trains and runs models **locally** so patient data never leaves the hospital/university network.

**Innovation:** direct WSI analysis + Iranian native multimodal dataset + clinical fusion + privacy-preserving local stack.  
**Output:** a maintainable CDS prototype that raises diagnostic accuracy/speed, lowers screening cost, and reduces dependence on foreign cloud tools.

## 3. Keywords (max 5)
Digital Pap smear · Artificial intelligence · Machine learning · Cervical cancer · Digital pathology

## 4. Duration
**12 months**

## 5. Technology Domain
- ☑ Medical devices, pharma & health  
- ☑ ICT  

## 6. Result Application
- ☑ Product prototype  
- ☑ Dataset creation  

## 7. Problem Definition (≤ 500 words)
Pap smear interpretation is specialist-dependent, high-volume, error-prone, and unevenly available across regions. Global AI solutions usually:
- use public / non-local datasets,
- ignore clinical metadata,
- run on foreign clouds (privacy, cost, sovereignty risks),
- and do not fit Iranian clinical conditions.

Iran lacks a **native, multimodal, locally trainable** Pap smear CDS. This project closes that gap with a compact on-prem stack, a standard native dataset, and measurable clinical metrics (accuracy, sensitivity, usability).

## 8. Project Objectives
1. Design/implement intelligent algorithms for digital Pap smear image analysis.  
2. Develop normal / abnormal cervical cell classification models.  
3. Fuse image features with patient clinical information.  
4. Build a standardized native Pap smear + clinical dataset.  
5. Evaluate accuracy, sensitivity, and clinical usability.

## 9. Literature Snapshot (APA)
- Litjens, G., et al. (2017). A survey on deep learning in medical image analysis. *Medical Image Analysis*.  
- Zhang, L., et al. (2020). Automated cervical cancer screening using deep learning. *IEEE Access*.  
- Komura, D., & Ishikawa, S. (2018). Machine learning methods for histopathological image analysis. *Computational and Structural Biotechnology Journal*.

Global work shows deep learning can match experts on abnormal-cell detection, but most systems are image-only and non-local. Iranian work is mostly small/cropped datasets—no full native multimodal CDS reported.

## 10. Necessity (≤ 200 words)
Early detection saves lives and cost. Smart CDS offsets pathologist shortage. Local training protects PHI, cuts foreign-vendor lock-in, and supports digital-health policy and equitable screening access.

## 11. Public Explanation (≥ 200 words)
The system helps doctors find suspicious cervical cells earlier. It reads digital Pap smear slides, flags abnormal cells, and also considers age and medical history—like a doctor would. Everything runs on hospital/university computers so patient data stays private. Result: faster, more consistent screening, even in remote areas, plus a reusable Iranian dataset for future research.

## 12. Challenges & Mitigations (≤ 200 words)
| Challenge | Mitigation |
|---|---|
| Lack of native data | Partner with provincial medical university / hospitals |
| Slide quality variation | Standard protocols + augmentation + robust models |
| Patient privacy | De-identification + fully local training/inference |
| Class imbalance | Weighted loss, oversampling, focused hard-negative mining |
| Maintainability | Compact 5-component stack (see Part C) |

## 13. Deliverables
- Scientific papers  
- Native reusable dataset  
- On-prem CDS prototype  
- Path toward patent / commercialization  
- Better diagnostic workflow quality  

## 14. Knowledge Acquisition
☑ Internal R&D (not reverse engineering / tech transfer)

## 15. Market
☑ National  ☑ International  

**Competitors:** no local entrant yet in this exact multimodal on-prem niche.  
**Similar products:** exist abroad (foreign goods), not as a privacy-first Iranian native stack.

## 16. Commercialization Readiness
| Item | Answer |
|---|---|
| Final outcome | Technology product |
| Commercialization intent | By the team |
| End beneficiaries | All women (screening population) |
| Institutional users | Public & private hospitals / clinics |
| Needed support | Medical university partnership for data collection |
| Next step after project | National commercialization / rollout |

---

# Part B — What the Document Means (Our Understanding)

The source form is not only a research idea; it defines a **national digital-health product path**.

### B1. Core intent
Build a **decision-support system** (not a fully autonomous diagnosis robot) that helps pathologists screen Pap smears faster and more consistently.

### B2. What “success” looks like
| Objective | Success signal |
|---|---|
| Algorithms for digital Pap images | Working WSI→tile→cell pipeline |
| Normal/abnormal classification | Validated model metrics (Acc / Sens / Spec / AUC) |
| Clinical fusion | Model that uses age, history, symptoms with images |
| Native dataset | Versioned, de-identified, documented corpus |
| Clinical usability | Pathologist pilot feedback + prototype CDS UI |

### B3. Non-negotiables implied by the form
1. **Whole-slide analysis** (no manual crop workflow).  
2. **Multimodal** (image + clinical).  
3. **Native Iranian data**.  
4. **Prototype product + dataset** (not only a paper).  
5. **Internal R&D** and path to commercialization.  
6. **Privacy** (explicit challenge in the form → local training).

### B4. What we will *not* do in v1
- Replace the pathologist.  
- Depend on Roboflow / foreign cloud labeling for PHI.  
- Build a bloated multi-tool platform.  
- Ship without evaluation on accuracy & sensitivity.

---

# Part C — The Way (Implementation Path)

```text
1. Partner hospitals → collect WSI + clinical fields (ethics + consent)
2. De-identify & store on local NAS
3. Pathologists annotate in QuPath
4. Train models locally on GPU workstation (PyTorch)
5. Log metrics in MLflow; iterate
6. Serve CDS prototype via FastAPI on hospital LAN
7. Pilot with pathologists → refine → prepare commercialization package
```

### C1. Compact maintainable stack (chosen)

| # | Component | Role |
|---|---|---|
| 1 | **OpenSlide** | Read WSI (`.svs`, `.ndpi`, `.tiff`) and cut tiles |
| 2 | **QuPath** | Pathologist annotation → native labeled dataset |
| 3 | **PyTorch** | Segmentation + classification + clinical fusion |
| 4 | **MLflow** | Experiment tracking, metrics, model versions |
| 5 | **FastAPI** | Local CDS API / prototype for hospitals |

**Hardware (one box is enough for year 1):**
- GPU workstation: RTX 4090 (24GB) or equivalent  
- Local NAS: 10–40 TB for WSI archive + backups  
- Hospital LAN only (no required internet for inference)

---

# Part D — Detailed Stack Chart

## D1. End-to-end architecture

```text
┌──────────────────────────────────────────────────────────────────────────┐
│                     HOSPITAL / UNIVERSITY PRIVATE LAN                     │
│                         (patient data never leaves)                       │
│                                                                          │
│  ┌─────────────┐   ┌──────────────┐   ┌────────────────────────────────┐ │
│  │ Slide scanner│──▶│  WSI store   │──▶│ OpenSlide                      │ │
│  │ (digital Pap)│   │  (NAS/HDD)   │   │ • read gigapixel slides        │ │
│  └─────────────┘   └──────────────┘   │ • tissue / tile extraction     │ │
│                                        │ • QC filters (blur/empty)      │ │
│                                        └─────────────┬──────────────────┘ │
│                                                      │ tiles + coords      │
│                                                      ▼                     │
│  ┌─────────────────────────────────────────────────────────────────────┐ │
│  │ QuPath (pathologist workstation)                                    │ │
│  │ • draw nuclei / cells                                               │ │
│  │ • label: normal / abnormal (+ optional Bethesda-style grades)       │ │
│  │ • export GeoJSON / mask + CSV clinical link IDs                     │ │
│  └───────────────────────────────┬─────────────────────────────────────┘ │
│                                  │ labeled dataset                         │
│                                  ▼                                         │
│  ┌─────────────────────────────────────────────────────────────────────┐ │
│  │ Dataset vault (local)                                               │ │
│  │ • /wsi  /tiles  /masks  /clinical.parquet  /splits.yaml             │ │
│  │ • de-identified IDs only                                            │ │
│  │ • encrypted disk + access control                                   │ │
│  └───────────────────────────────┬─────────────────────────────────────┘ │
│                                  │                                         │
│                                  ▼                                         │
│  ┌─────────────────────────────────────────────────────────────────────┐ │
│  │ PyTorch training box (GPU)                                          │ │
│  │  A) Segmentation model  → find cells/nuclei                         │ │
│  │  B) Classification model → normal vs abnormal                       │ │
│  │  C) Clinical fusion head → age, history, symptoms → patient risk    │ │
│  │  D) Optional MIL aggregator → slide-level score                     │ │
│  └───────────────────────────────┬─────────────────────────────────────┘ │
│                                  │                                         │
│                                  ▼                                         │
│  ┌──────────────────────┐      ┌───────────────────────────────────────┐ │
│  │ MLflow               │      │ FastAPI CDS prototype                 │ │
│  │ • Acc / Sens / Spec  │      │ • upload/select case                  │ │
│  │ • AUC / F1           │      │ • show flagged cells + risk score     │ │
│  │ • model registry     │◀────▶│ • pathologist confirm / override      │ │
│  └──────────────────────┘      │ • audit log                           │ │
│                                 └───────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────────────────┘
```

## D2. Component detail sheet

| Layer | Tool | Inputs | Outputs | Why kept |
|---|---|---|---|---|
| WSI I/O | OpenSlide | Scanner files | Tiles + coordinates | Standard, free, local |
| Annotation | QuPath | WSI | Masks / GeoJSON labels | Pathologist-native UI |
| Vision + fusion | PyTorch | Tiles, masks, clinical table | Models + scores | Full control, multimodal |
| Tracking | MLflow | Runs, metrics, weights | Reproducible registry | Lightweight vs heavy MLOps |
| Serving | FastAPI | Model + case request | JSON / simple review UI | Easy hospital deploy |

## D3. Explicitly excluded (to stay maintainable)
Roboflow cloud, Label Studio/CVAT (unless QuPath bottlenecks), Kubernetes, multi-cloud MLOps, heavy front-end frameworks in v1, separate feature stores.

---

# Part E — How We Train the Models (Local)

## E1. Training stages

| Stage | Model | Labels needed | Output |
|---|---|---|---|
| **1. Segmentation** | U-Net / Cellpose-style / custom PyTorch seg | Nucleus/cell outlines from QuPath | Cell masks |
| **2. Cell classification** | CNN/ViT classifier (timm backbones) | Normal / abnormal per cell | Cell probabilities |
| **3. Clinical fusion** | Late fusion MLP / gated fusion | Clinical fields + image embedding | Case-level risk |
| **4. (Optional) Slide MIL** | Attention-MIL | Slide/case label | Slide score + heat highlights |

## E2. Local training procedure
1. Split data by **patient ID** (train/val/test) to avoid leakage.  
2. Augment tiles (stain jitter, flip, rotate, blur) for scanner/stain diversity.  
3. Train on local GPU; checkpoint best validation sensitivity (screening priority).  
4. Calibrate thresholds with pathologists (high recall first).  
5. Register best weights in MLflow; promote only after hold-out test.  
6. Package ONNX/TorchScript for FastAPI inference **on the same LAN**.

## E3. Metrics tied to form objective #5
- Accuracy, Sensitivity (Recall), Specificity, Precision, F1, AUC  
- Pathologist time saved / agreement rate (usability pilot)  
- Fail cases reviewed weekly

## E4. Data volume guidance (year 1)
| Asset | Target (practical) |
|---|---|
| WSI cases | 500–2,000 (start), grow later |
| Annotated cells | tens of thousands |
| Clinical fields | age, prior Rx, symptoms, HPV if available, outcome label |

---

# Part F — Why Train Locally (vs Online Platforms)

## F1. Decision
**Train and infer on-prem.** Online platforms (e.g. Roboflow) may help demos on **public** cytology images only—not Iranian patient PHI.

## F2. Comparison table

| Criterion | Local stack (recommended) | Online (e.g. Roboflow) |
|---|---|---|
| **Patient privacy** | Data stays in hospital LAN | Free plan = public; paid still vendor cloud |
| **Fits WSI + clinical fusion** | Yes (custom PyTorch) | Weak (patch CV workflows) |
| **Cost model** | One-time GPU + storage | Subscription + credits forever |
| **Training iterations** | Unlimited locally | Credits: ~1 credit / 30 min GPU |
| **Self-hosted inference** | Free after setup | Still consumes credits if platform-tied |
| **Sovereignty / sanctions risk** | Controlled | Foreign SaaS dependency |
| **Native dataset ownership** | Full | Locked into vendor formats/process |
| **Maintainability** | 5 tools | Vendor roadmap + billing complexity |
| **Hospital IT acceptance** | Higher (air-gap capable) | Often blocked for PHI |

## F3. Roboflow price snapshot (indicative)
| Plan | Price | Notes |
|---|---|---|
| Public | Free | Data shared publicly — **unusable for PHI** |
| Core | ~$79–99 / month | Private workspace; limited included credits |
| Enterprise | Custom | Needed for serious compliance; still not a WSI multimodal stack |

**Credit examples:** 1 credit ≈ 30 min GPU training; storage/inference also burn credits—even for some self-hosted usage of platform models.

## F4. Year-1 cost sketch
| Approach | Approx. year-1 | Fit to proposal |
|---|---|---|
| Local GPU + NAS | **$2k–$6k** one-time (+ power) | Strong |
| Roboflow Core + overages | **$2.5k–$8k+** recurring | Weak for WSI/PHI/fusion |
| Roboflow Enterprise | Often much higher | Overkill / misaligned |

**Conclusion:** Local training is better for **privacy, cost control, technical fit, and national technology goals** stated in the form.

---

# Part G — Data Storage Design

## G1. What we store
| Bucket | Content | Format |
|---|---|---|
| Raw WSI | Original scanner slides | `.svs` / `.ndpi` / `.tiff` |
| Tiles | Extracted patches + coords | `.png`/`.jpg` + `tiles.parquet` |
| Annotations | Nucleus masks / polygons | GeoJSON / PNG masks |
| Clinical | De-identified tabular fields | Parquet / CSV |
| Models | Checkpoints + ONNX | MLflow registry |
| Audit | Who viewed/confirmed cases | Append-only logs |

## G2. Privacy rules
1. **No national ID / name / phone** in training folders.  
2. Use irreversible study IDs (`CASE_000123`).  
3. Keep re-identification map **offline**, access-controlled, separate disk.  
4. Encrypt NAS volumes; role-based access (pathologist / ML / admin).  
5. No upload of PHI to cloud labeling, chatbots, or SaaS trainers.  
6. Ethics approval + informed consent / waiver per university protocol.

## G3. Folder layout (simple & maintainable)

```text
/data/
  raw_wsi/
  tiles/
  annotations/
  clinical/
    clinical.parquet
  splits/
    train.txt  val.txt  test.txt
  models/          # or MLflow artifact store
  exports/         # de-identified shares for partners
```

## G4. Backup
- Daily incremental backup to second local disk/NAS  
- Weekly offline cold copy  
- Test restore once per quarter  

---

# Part H — Sharing with Hospitals (Without Breaking Privacy)

## H1. Sharing principle
Share **software + de-identified evaluation packages + deployment guides**, not raw identifiable patient archives.

## H2. Hospital collaboration modes

| Mode | What hospital gets | What hospital sends back |
|---|---|---|
| **Data partner** | Annotation protocol + QuPath project template | De-identified WSI + clinical fields via secure disk/VPN |
| **Pilot site** | FastAPI CDS appliance (LAN install) | Usage metrics + pathologist feedback |
| **Model update loop** | New model weights (MLflow package) | Optional federated/aggregated metrics only |
| **Multi-center study** | Common schema + SOPs | Harmonized de-identified exports |

## H3. Practical sharing channels
1. **Encrypted external drive** handoff (preferred for large WSI).  
2. **Hospital-to-university VPN** for smaller batches.  
3. **On-site deployment**: engineers install CDS on hospital GPU/CPU box; data never copied out.  
4. **Federated-style updates (later):** train at site, share weights/gradients—not images.

## H4. What a hospital deployment package contains
- FastAPI service + model weights  
- OpenSlide runtime  
- Simple review UI (flagged cells + risk + override)  
- Admin guide (users, backup, GPU drivers)  
- Clinical SOP (who confirms AI suggestions)

## H5. Governance checklist before any share
- [ ] Ethics / IRB approval  
- [ ] Data-sharing agreement (DSA)  
- [ ] De-identification verified  
- [ ] Access roles defined  
- [ ] Audit logging enabled  
- [ ] Incident response contact named  

---

# Part I — 12-Month Timeline

```text
Month  1  2  3  4  5  6  7  8  9 10 11 12
       |--|--|--|--|--|--|--|--|--|--|--|--|
Ethics & partners ████
Storage + OpenSlide pipeline    ████████
QuPath annotation protocol         ████████████
Native dataset v1                     ████████
Seg model training                       ████████
Cell classifier                          ████████
Clinical fusion                             ████████
MLflow eval / iteration                        ████████
FastAPI CDS prototype                             ████████
Hospital pilot + usability                           ████████
Docs, paper, commercialization pack                     ████
```

### Phase plan

| Phase | Months | Work | Exit criteria |
|---|---|---|---|
| **P0 Setup** | 1–2 | Ethics, hospital MoU, hardware, folder schema | Approvals + NAS online |
| **P1 Data** | 2–5 | WSI ingest, QuPath labeling, clinical table | Dataset v1 frozen |
| **P2 Models** | 4–8 | Seg + classify + fusion training locally | Hold-out metrics meet pilot bar |
| **P3 CDS** | 7–10 | FastAPI prototype, review UI, audit log | Deployable LAN build |
| **P4 Pilot** | 9–11 | 1–2 hospital pilots, pathologist feedback | Usability report |
| **P5 Close** | 11–12 | Paper, dataset card, commercialization plan | Form deliverables complete |

### Milestone ↔ form objectives
| Form objective | Milestone month |
|---|---|
| Algorithms for Pap images | M6 |
| Normal/abnormal models | M8 |
| Clinical fusion | M9 |
| Native dataset | M5 (v1), M11 (v1.1) |
| Acc/Sens/usability eval | M10–M12 |

---

# Part J — Risks & Controls

| Risk | Impact | Control |
|---|---|---|
| Slow annotation | Delays dataset | Start with triage labels; add fine labels later |
| Scanner stain shift | Drop in accuracy | Color augmentation; multi-site samples |
| Class imbalance | Low sensitivity | Sensitivity-first threshold; focal/weighted loss |
| Hospital IT blockers | No pilot | Provide offline appliance image |
| Scope creep | Unmaintainable stack | Keep the 5-tool rule |

---

# Part K — Final Recommendation

1. **Follow the form’s five objectives exactly.**  
2. **Use the compact local stack:** OpenSlide → QuPath → PyTorch → MLflow → FastAPI.  
3. **Train only on-prem** for privacy, cost, and technical fit; avoid online SaaS for PHI.  
4. **Store de-identified data on local NAS** with strict access and backups.  
5. **Share with hospitals via secure handoff / on-site deploy / later weight sharing**—never raw identifiable dumps.  
6. **Deliver in 12 months:** native dataset + evaluated models + LAN CDS prototype + commercialization package.

```text
WSI (local) → OpenSlide → QuPath labels → PyTorch train (local GPU)
        → MLflow metrics → FastAPI CDS on hospital LAN → pathologist confirms
```

This is the shortest maintainable path that still fully satisfies the uploaded technology-development request form.
