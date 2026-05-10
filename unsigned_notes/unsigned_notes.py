"""
Open Dental - Unsigned Notes Daily Email Reminder
--------------------------------------------------
Sends each provider a daily email listing completed appointments
from the last 30 days that have NO signed procedure note.

Business Rules:
- Only completed appointments (AptStatus = 2)
- Rolling 30-day lookback
- Skip appointments with no procedures
- Appointment is OK if ANY one procedure has a signed note (mouse or Topaz pad)
- Appointment is OK if a signed ~GRP~ group note exists for same patient+date
- Covers both doctor and hygiene appointments
- Runs daily at 10AM via Windows Task Scheduler
"""

import mysql.connector
import smtplib
import logging
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta

# =============================================================================
# CONFIG — loaded from office_config.py
# To deploy at a new office, only edit office_config.py
# =============================================================================

import sys
sys.path.insert(0, r"C:\dental_automation\config")

try:
    from office_config import (
        OFFICE, DB_CONFIG, GMAIL_CONFIG,
        PROVIDER_EMAILS, FOLDERS,
        UNSIGNED_NOTES_SETTINGS
    )
except ImportError as e:
    print(f"ERROR: Could not load office_config.py: {e}")
    print("Make sure office_config.py exists in C:\\dental_automation\\config\\")
    sys.exit(1)

PRACTICE_NAME        = OFFICE["name"]
LOG_FILE             = os.path.join(FOLDERS["unsigned_notes"], "unsigned_notes.log")
LOOKBACK_DAYS        = UNSIGNED_NOTES_SETTINGS["lookback_days"]
PERIO_LOOKAHEAD_DAYS = UNSIGNED_NOTES_SETTINGS["perio_lookahead_days"]
NOTES_RETENTION_DAYS = UNSIGNED_NOTES_SETTINGS["notes_retention_days"]
PROVIDER_NOTES_DIR   = FOLDERS["provider_notes"]


NOTE_TEMPLATES = {
    "D0120": (
        "PERIODIC ORAL EVALUATION\n\n"
        "Pt presents today for periodic oral evaluation.\n"
        "Med hx reviewed with no significant changes noted.\n"
        "Radiographs taken and reviewed - [specify views]\n"
        "Intraoral and extraoral examination completed.\n"
        "Periodontal evaluation - [WNL / findings]\n"
        "Radiographic findings - [describe]\n"
        "Existing restorations evaluated - [stable / concerns noted]\n"
        "Discussed findings with patient and reviewed treatment options.\n"
        "Discussed risks and benefits of treatment vs. no treatment. Patient questions answered.\n\n"
        "Plan: [list treatment]\n\n"
        "OH - [Good / Fair / Poor]\n"
        "OHI provided and reviewed.\n\n"
        "NV - [next visit]"
    ),
    "D0150": (
        "COMPREHENSIVE ORAL EVALUATION\n\n"
        "Pt presents today for comprehensive oral evaluation.\n"
        "Med hx reviewed and documented. [Note any significant medical history]\n"
        "Radiographs taken - [specify: panoramic, bitewings, periapicals]\n"
        "Full intraoral and extraoral examination completed.\n"
        "Periodontal screening completed - [WNL / findings]\n"
        "Soft tissue evaluation - [WNL / findings]\n"
        "Occlusal evaluation - [WNL / findings]\n"
        "Radiographic findings reviewed and discussed with patient.\n"
        "Discussed diagnosis and treatment options.\n"
        "Discussed risks and benefits of treatment vs. no treatment. Patient questions answered.\n"
        "Patient consent obtained for proposed treatment plan.\n\n"
        "Plan: [list treatment]\n\n"
        "OH - [Good / Fair / Poor]\n"
        "OHI provided and reviewed.\n\n"
        "NV - [next visit]"
    ),
    "D0140": (
        "LIMITED ORAL EVALUATION\n\n"
        "Pt presents today for limited oral evaluation - [chief complaint / tooth #]\n"
        "Med hx reviewed with no significant changes noted.\n"
        "Radiographs taken - [specify views]\n"
        "Radiographic and intraoral evaluation reveals - [findings]\n"
        "Recommended tx - [diagnosis and proposed treatment]\n\n"
        "Pt was informed about risks vs. benefits of the proposed treatment, alternative treatment options,\n"
        "and the option of no treatment. Patient was informed they may refuse treatment at any point.\n\n"
        "Tx done today - [if any]\n\n"
        "NV - [next visit]"
    ),
    "D1110": (
        "ADULT PROPHYLAXIS\n\n"
        "Pt presents today for adult prophylaxis.\n"
        "Med hx reviewed with no significant changes noted.\n"
        "Periodontal evaluation completed - [WNL / findings]\n"
        "Supragingival and subgingival calculus removal completed with piezo and hand instrumentation.\n"
        "Teeth polished with prophy cup on slow-speed handpiece using prophy paste.\n"
        "Interproximal areas flossed.\n"
        "Fluoride applied - [if applicable]\n"
        "OHI provided - proper brushing and flossing techniques demonstrated and reviewed.\n\n"
        "Radiographic findings reviewed - [if applicable]\n"
        "Discussed findings with patient and reviewed treatment plan.\n\n"
        "NV - [next visit / recall interval]"
    ),
    "D1120": (
        "CHILD PROPHYLAXIS\n\n"
        "Pt presents today for child prophylaxis.\n"
        "Med hx reviewed with no significant changes noted.\n"
        "Periodontal evaluation completed - [WNL / findings]\n"
        "Supragingival calculus removal completed.\n"
        "Teeth polished with prophy cup using prophy paste.\n"
        "Interproximal areas flossed.\n"
        "Fluoride applied - [if applicable]\n"
        "OHI provided - brushing and flossing techniques demonstrated and reviewed with patient and parent/guardian.\n\n"
        "Discussed findings with parent/guardian and reviewed treatment plan.\n\n"
        "NV - [next visit / recall interval]"
    ),
    "D4341": (
        "PERIODONTAL SCALING AND ROOT PLANING\n\n"
        "Pt presents today for SRP - [specify quadrant(s)]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "SRP completed with piezo ultrasonic and hand instrumentation.\n"
        "Root surfaces debrided and smoothed.\n"
        "The area was irrigated with Peridex (chlorhexidine 0.12%).\n\n"
        "Probing depths pre-treatment: [record or reference periodontal chart]\n"
        "Patient tolerated procedure well.\n\n"
        "Post-op instructions provided:\n"
        "- Mild soreness expected for 1-2 days\n"
        "- Rinse with warm salt water 2-3x daily\n"
        "- Avoid hard/crunchy foods for 24 hours\n"
        "- Continue prescribed oral hygiene regimen\n\n"
        "NV - [re-evaluation / remaining quads]"
    ),
    "D4342": (
        "PERIODONTAL SCALING AND ROOT PLANING (1-3 TEETH)\n\n"
        "Pt presents today for SRP - [specify teeth]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "SRP completed with piezo ultrasonic and hand instrumentation on teeth [#].\n"
        "Root surfaces debrided and smoothed.\n"
        "The area was irrigated with Peridex (chlorhexidine 0.12%).\n\n"
        "Patient tolerated procedure well.\n\n"
        "Post-op instructions provided:\n"
        "- Mild soreness expected for 1-2 days\n"
        "- Rinse with warm salt water 2-3x daily\n"
        "- Avoid hard/crunchy foods for 24 hours\n\n"
        "NV - [re-evaluation]"
    ),
    "D2391": (
        "COMPOSITE RESTORATION - ONE SURFACE POSTERIOR\n\n"
        "Pt presents today for composite restoration - [tooth # / surface]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Rubber dam / isolation placed.\n"
        "Caries removed with handpiece and spoon excavator.\n"
        "Cavity prepared and etched. Bonding agent applied and light cured.\n"
        "Composite resin placed incrementally and light cured.\n"
        "Restoration contoured, finished, and polished.\n"
        "Occlusion checked and adjusted as needed.\n\n"
        "Patient tolerated procedure well.\n\n"
        "NV - [next visit]"
    ),
    "D2392": (
        "COMPOSITE RESTORATION - TWO SURFACES POSTERIOR\n\n"
        "Pt presents today for composite restoration - [tooth # / surfaces]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Rubber dam / isolation placed.\n"
        "Caries removed with handpiece and spoon excavator.\n"
        "Cavity prepared and etched. Bonding agent applied and light cured.\n"
        "Matrix band and wedge placed for interproximal contact.\n"
        "Composite resin placed incrementally and light cured.\n"
        "Restoration contoured, finished, and polished.\n"
        "Occlusion checked and adjusted as needed.\n\n"
        "Patient tolerated procedure well.\n\n"
        "NV - [next visit]"
    ),
    "D2393": (
        "COMPOSITE RESTORATION - THREE SURFACES POSTERIOR\n\n"
        "Pt presents today for composite restoration - [tooth # / surfaces]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Rubber dam / isolation placed.\n"
        "Caries removed with handpiece and spoon excavator.\n"
        "Cavity prepared and etched. Bonding agent applied and light cured.\n"
        "Matrix band and wedge placed for interproximal contact.\n"
        "Composite resin placed incrementally and light cured.\n"
        "Restoration contoured, finished, and polished.\n"
        "Occlusion checked and adjusted as needed.\n\n"
        "Patient tolerated procedure well.\n\n"
        "NV - [next visit]"
    ),
    "D2331": (
        "COMPOSITE RESTORATION - TWO SURFACES ANTERIOR\n\n"
        "Pt presents today for composite restoration - [tooth # / surfaces]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Isolation placed.\n"
        "Caries/defect removed with handpiece and hand instruments.\n"
        "Cavity prepared, etched, bonding agent applied and light cured.\n"
        "Shade selected: [shade]\n"
        "Composite resin placed incrementally and light cured.\n"
        "Restoration contoured, finished, and polished to natural tooth appearance.\n"
        "Occlusion and contacts checked and adjusted as needed.\n\n"
        "Patient tolerated procedure well. Patient satisfied with esthetics.\n\n"
        "NV - [next visit]"
    ),
    "D2332": (
        "COMPOSITE RESTORATION - THREE SURFACES ANTERIOR\n\n"
        "Pt presents today for composite restoration - [tooth # / surfaces]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Isolation placed.\n"
        "Caries/defect removed with handpiece and hand instruments.\n"
        "Cavity prepared, etched, bonding agent applied and light cured.\n"
        "Shade selected: [shade]\n"
        "Composite resin placed incrementally and light cured.\n"
        "Restoration contoured, finished, and polished to natural tooth appearance.\n"
        "Occlusion and contacts checked and adjusted as needed.\n\n"
        "Patient tolerated procedure well. Patient satisfied with esthetics.\n\n"
        "NV - [next visit]"
    ),
    "D2335": (
        "COMPOSITE RESTORATION - FOUR OR MORE SURFACES / INCISAL ANGLE ANTERIOR\n\n"
        "Pt presents today for composite restoration involving incisal angle - [tooth # / surfaces]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Isolation placed.\n"
        "Existing caries/fracture/defect removed.\n"
        "Cavity prepared, etched, bonding agent applied and light cured.\n"
        "Shade selected: [shade]\n"
        "Composite resin placed incrementally with attention to incisal edge morphology. Light cured.\n"
        "Restoration contoured, finished, and polished to natural tooth appearance.\n"
        "Occlusion and incisal contacts checked and adjusted as needed.\n\n"
        "Patient tolerated procedure well. Patient satisfied with esthetics.\n\n"
        "NV - [next visit]"
    ),
    "D2740": (
        "CROWN - PORCELAIN/CERAMIC DELIVERY\n\n"
        "Pt presents today for porcelain/ceramic crown delivery - [tooth #]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Temporary crown removed. Tooth and preparation cleaned.\n"
        "Permanent crown tried in - fit, contour, and contacts verified.\n"
        "Shade verified with patient approval.\n"
        "Occlusion checked in centric and lateral excursions.\n"
        "Adjustments made as needed.\n"
        "Crown cemented with [cement type].\n"
        "Excess cement removed. Final occlusion verified.\n\n"
        "Patient tolerated procedure well. Patient satisfied with esthetics.\n\n"
        "NV - [next visit as needed]"
    ),
    "D2700": (
        "CROWN PREPARATION\n\n"
        "Pt presents today for crown preparation - [tooth #]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Existing restoration/caries removed.\n"
        "Tooth prepared for full coverage crown.\n"
        "Retraction cord placed. Final impression taken with [material].\n"
        "Opposing arch impression taken.\n"
        "Bite registration recorded.\n"
        "Shade selected: [shade]\n"
        "Temporary crown fabricated and cemented with temporary cement.\n"
        "Occlusion verified and adjusted.\n\n"
        "Patient tolerated procedure well.\n"
        "Lab prescription sent to [lab name].\n\n"
        "NV - Crown delivery [estimated date]"
    ),
    "D7140": (
        "SIMPLE EXTRACTION\n\n"
        "Pt presents today for extraction of tooth #[__]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Risks, benefits, and alternatives to extraction discussed with patient.\n"
        "Informed consent obtained.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Tooth elevated and extracted with forceps without complication.\n"
        "Socket irrigated. Hemostasis achieved with gauze pressure.\n"
        "No complications encountered.\n\n"
        "Post-op instructions provided (verbal and written):\n"
        "- Bite on gauze for 30-45 minutes\n"
        "- Avoid smoking, spitting, straws for 72 hours\n"
        "- Soft diet for 24-48 hours\n"
        "- OTC pain management as directed\n\n"
        "Rx: [if prescribed]\n\n"
        "NV - [post-op / as needed]"
    ),
    "D7210": (
        "SURGICAL EXTRACTION\n\n"
        "Pt presents today for surgical extraction of tooth #[__]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Risks, benefits, and alternatives discussed with patient.\n"
        "Informed consent obtained.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Mucoperiosteal flap reflected. Bone removal performed as necessary.\n"
        "Tooth sectioned [if applicable]. Tooth elevated and extracted.\n"
        "Socket debrided and irrigated.\n"
        "[Bone graft placed - if applicable]\n"
        "[Membrane placed - if applicable]\n"
        "Flap repositioned and sutured with [suture type].\n"
        "Hemostasis achieved.\n\n"
        "No complications encountered. Patient tolerated procedure well.\n\n"
        "Post-op instructions provided (verbal and written):\n"
        "- Bite on gauze for 30-45 minutes\n"
        "- Avoid smoking, spitting, straws for 72 hours\n"
        "- Soft diet recommended\n"
        "- Ice pack 20 min on/off for first 24 hours\n"
        "- OTC/prescribed pain management as directed\n\n"
        "Rx: [Amoxicillin 500mg / Ibuprofen 800mg / other]\n\n"
        "Suture removal scheduled: [date]\n"
        "NV - Post-op [date]"
    ),
    "D7953": (
        "BONE REPLACEMENT GRAFT - RIDGE PRESERVATION\n\n"
        "Pt presents today for bone graft / ridge preservation at site #[__]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Following extraction at site #[__], ridge preservation bone graft placed.\n"
        "Graft material: [specify - allograft / xenograft / synthetic]\n"
        "Membrane placed: [yes/no - specify type if yes]\n"
        "Socket irrigated and graft material packed.\n"
        "Flap sutured over graft site.\n\n"
        "Patient tolerated procedure well.\n"
        "Post-op instructions provided.\n\n"
        "Rx: [if prescribed]\n\n"
        "NV - Post-op / implant planning [timeframe]"
    ),
    "D3330": (
        "ENDODONTIC THERAPY - MOLAR\n\n"
        "Pt presents today for root canal therapy - tooth #[__]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Diagnosis: [Pulpal necrosis / Irreversible pulpitis / Symptomatic apical abscess]\n"
        "Radiographic findings: [radiolucency at apex / thickened PDL / other]\n\n"
        "Risks, benefits, and alternatives (including extraction) discussed with patient.\n"
        "Informed consent obtained.\n\n"
        "Topical anesthesia applied. [__] carpules of 2% lidocaine with 1:100,000 epi administered.\n"
        "Isolation with rubber dam.\n"
        "Access opening made. Pulp chamber entered. Pulp tissue removed.\n\n"
        "Canals located and negotiated:\n"
        "MB: [__] mm  |  ML: [__] mm  |  DB: [__] mm  |  DL: [__] mm\n\n"
        "Canals shaped with rotary instrumentation.\n"
        "Irrigation: NaOCl, EDTA, saline, Peridex.\n"
        "Canals dried with paper points.\n"
        "Canals obturated with gutta percha and [sealer]. Working lengths confirmed radiographically.\n"
        "Access closed with [buildup material]. Occlusion checked and adjusted.\n\n"
        "Post-op instructions provided.\n"
        "Rx: [Amoxicillin 500mg / Ibuprofen 800mg]\n\n"
        "Permanent crown recommended.\n"
        "NV - Crown preparation"
    ),
    "D6010": (
        "IMPLANT PLACEMENT\n\n"
        "Pt presents today for implant placement at site #[__]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Risks, benefits, and alternatives discussed with patient.\n"
        "Informed consent obtained.\n\n"
        "Topical anesthesia applied at injection sites.\n"
        "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.\n\n"
        "Crestal incision made. Mucoperiosteal flap reflected.\n"
        "Osteotomy prepared per implant protocol.\n"
        "Implant placed: [brand / size] at [__] Ncm torque.\n"
        "Cover screw / healing abutment placed.\n"
        "Flap sutured with [suture type].\n"
        "Post-op radiograph taken to verify implant position.\n\n"
        "Patient tolerated procedure well.\n"
        "Post-op instructions provided.\n"
        "Rx: [Amoxicillin 500mg / Ibuprofen 800mg / Peridex rinse]\n\n"
        "Suture removal: [date]\n"
        "Integration period: [3-6 months]\n"
        "NV - Implant uncovery / crown"
    ),
    "D5201": (
        "COMPLETE DENTURE - IMPRESSION\n\n"
        "Pt presents today for complete denture impressions.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Preliminary/final impressions taken with [stock/custom] trays and [alginate/PVS].\n"
        "Upper and lower arch impressions evaluated for accuracy and completeness.\n"
        "Impressions poured and sent to lab.\n\n"
        "NV - Bite registration, wax rims, and shade selection"
    ),
    "D5202": (
        "DENTURE - WAX RIMS / BITE REGISTRATION\n\n"
        "Pt presents today for wax rim try-in and bite registration.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Custom trays with bite rims tried in.\n"
        "Adjustments made to intaglio surface and buccal flanges of upper and lower custom tray.\n"
        "Upper and lower rims adjusted to achieve adequate lip fullness and smile line.\n"
        "Plane of occlusion verified. Midline marked.\n"
        "Vertical dimension of occlusion established and recorded.\n"
        "Bite registration recorded.\n"
        "Shade and mold selected: [shade / mold]\n"
        "Patient satisfied with esthetics at this stage.\n\n"
        "NV - Denture try-in"
    ),
    "D5203": (
        "DENTURE TRY-IN\n\n"
        "Pt presents today for denture try-in.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Upper and lower wax try-in with teeth evaluated.\n"
        "Fit, retention, and stability assessed.\n"
        "Bite checked - centric occlusion verified.\n"
        "Midline confirmed. Phonetics evaluated.\n"
        "Esthetics verified - shade and tooth shape evaluated.\n"
        "Patient approved shape, shade, and overall appearance.\n\n"
        "NV - Denture delivery"
    ),
    "D5110": (
        "COMPLETE DENTURE DELIVERY - MAXILLARY\n\n"
        "Pt presents today for delivery of complete maxillary denture.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Denture tried in. Fit and retention verified.\n"
        "Flange extensions checked and adjusted as needed.\n"
        "Bite checked - centric occlusion verified and adjusted.\n"
        "Midline confirmed. Phonetics evaluated and found acceptable.\n"
        "Esthetics verified - patient satisfied.\n"
        "Patient instructed on insertion, removal, and denture care.\n"
        "Patient advised to leave dentures out at night and store in water.\n\n"
        "NV - Post-delivery adjustment as needed"
    ),
    "D5120": (
        "COMPLETE DENTURE DELIVERY - MANDIBULAR\n\n"
        "Pt presents today for delivery of complete mandibular denture.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Denture tried in. Fit, retention, and stability verified.\n"
        "Flange extensions checked and adjusted as needed.\n"
        "Bite checked - centric occlusion verified and adjusted.\n"
        "Midline confirmed. Phonetics evaluated and found acceptable.\n"
        "Esthetics verified - patient satisfied.\n"
        "Patient instructed on insertion, removal, and denture care.\n"
        "Patient advised to leave dentures out at night and store in water.\n\n"
        "NV - Post-delivery adjustment as needed"
    ),
    "D1351": (
        "SEALANT APPLICATION\n\n"
        "Pt presents today for sealant placement on tooth #[__]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Tooth isolated and dried.\n"
        "Enamel surfaces etched, rinsed, and dried.\n"
        "Sealant material applied and light cured.\n"
        "Sealant integrity and occlusion checked and adjusted as needed.\n\n"
        "Patient and parent/guardian instructed on sealant care and maintenance.\n\n"
        "NV - Routine recall"
    ),
    "D1208": (
        "TOPICAL FLUORIDE APPLICATION\n\n"
        "Pt presents today for topical fluoride application.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Fluoride varnish applied to all tooth surfaces following prophylaxis.\n"
        "Patient instructed not to eat or drink for 30 minutes.\n"
        "Patient advised to avoid hot foods/drinks for remainder of day.\n\n"
        "OHI reviewed.\n\n"
        "NV - Routine recall"
    ),
    "D9310": (
        "CONSULTATION\n\n"
        "Pt presents today for dental consultation.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Reason for consultation: [referral reason / chief complaint]\n"
        "Clinical and radiographic examination performed.\n"
        "Findings: [describe]\n"
        "Diagnosis: [describe]\n"
        "Treatment recommendations discussed with patient.\n"
        "Risks, benefits, and alternatives explained. Patient questions answered.\n"
        "[Referral back to treating dentist / treatment to be performed here]\n\n"
        "NV - [treatment / follow-up]"
    ),
    "D8090": (
        "ORTHODONTIC VISIT\n\n"
        "Pt presents today for orthodontic visit.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Treatment progress evaluated.\n"
        "[Adjustments made / wires changed / attachments placed or removed]\n"
        "Current tooth movement assessment: [describe]\n"
        "Patient compliance reviewed - [good / fair / poor]\n"
        "OHI reinforced - importance of oral hygiene with appliances reviewed.\n"
        "Estimated treatment timeline discussed.\n\n"
        "NV - [next orthodontic visit]"
    ),
    "D0220": (
        "PERIAPICAL RADIOGRAPH\n\n"
        "Periapical radiograph taken of tooth/area #[__]\n"
        "Radiograph reviewed.\n"
        "Findings: [describe]\n\n"
        "NV - [next visit]"
    ),
    "D0274": (
        "BITEWING RADIOGRAPHS\n\n"
        "Four bitewing radiographs taken.\n"
        "Radiographs reviewed.\n"
        "Interproximal caries evaluation: [WNL / caries noted at]\n"
        "Bone level evaluation: [WNL / bone loss noted]\n"
        "Existing restorations evaluated: [stable / concerns noted]\n\n"
        "Findings discussed with patient.\n\n"
        "NV - [next visit]"
    ),
    "D0330": (
        "PANORAMIC RADIOGRAPH\n\n"
        "Panoramic radiograph taken and reviewed.\n"
        "Findings: [describe - TMJ, bone levels, pathology, impacted teeth, etc.]\n"
        "All visible structures evaluated.\n"
        "Significant findings discussed with patient.\n\n"
        "NV - [next visit]"
    ),
    "POST": (
        "POST-OPERATIVE VISIT\n\n"
        "Pt presents today for post-operative visit following [procedure] on tooth #[__]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "The surgical/treatment site is healing well and within normal limits.\n"
        "Area irrigated with normal saline.\n"
        "Sutures / membrane removed at site #[__] - [if applicable]\n"
        "No signs of infection, dehiscence, or abnormal healing noted.\n"
        "Patient advised to continue rinsing with warm salt water.\n"
        "Patient reports [pain level / no pain / mild discomfort] - manageable with OTC medication.\n\n"
        "NV - [next visit as needed]"
    ),
    "Office Visit": (
        "OFFICE VISIT\n\n"
        "Pt presents today for office visit.\n"
        "Chief complaint: [describe reason for visit]\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Clinical findings: [describe examination findings]\n"
        "Assessment: [diagnosis / impression]\n\n"
        "Treatment rendered today: [describe]\n"
        "Patient informed of findings and treatment options.\n"
        "Risks, benefits, and alternatives discussed. Patient questions answered.\n\n"
        "NV - [next visit]"
    ),
    "D0082": (
        "VIVERA RETAINER IMPRESSION\n\n"
        "Pt presents today for Vivera retainer impression.\n"
        "Med hx reviewed with no significant changes noted.\n\n"
        "Upper and lower impressions taken for Vivera retainer fabrication.\n"
        "Impressions evaluated for accuracy.\n"
        "Impressions sent to Invisalign/Vivera lab.\n\n"
        "Patient instructed on retainer wear and care upon delivery.\n\n"
        "NV - Retainer delivery"
    ),
    "D0601": (
        "CARIES RISK ASSESSMENT\n\n"
        "Caries risk assessment completed.\n"
        "Risk level determined: [Low / Moderate / High]\n"
        "Contributing risk factors reviewed with patient: [diet, oral hygiene, fluoride exposure, etc.]\n"
        "Preventive recommendations discussed.\n"
        "Documentation completed.\n\n"
        "NV - [next visit]"
    ),
}

DEFAULT_NOTE_TEMPLATE = (
    "CLINICAL NOTE\n\n"
    "Pt presents today for [procedure description].\n"
    "Med hx reviewed with no significant changes noted.\n\n"
    "Clinical findings: [describe]\n"
    "Treatment rendered: [describe]\n"
    "Patient tolerated procedure well.\n\n"
    "Post-op instructions provided as applicable.\n\n"
    "NV - [next visit]"
)


# =============================================================================
# QUERIES
# =============================================================================

QUERY_UNSIGNED_APTS = """
SELECT
    a.AptNum,
    a.AptDateTime,
    a.ProvNum,
    a.ProvHyg,
    a.IsHygiene,
    CONCAT(pat.LName, ', ', pat.FName) AS PatientName,
    -- Check 1: ANY procedure in this appointment has a signed procnote
    -- Covers both mouse/digital signatures and Topaz pad signatures
    MAX(
        CASE
            WHEN (pn.Signature IS NOT NULL AND pn.Signature != '')
              OR pn.SigIsTopaz = 1 THEN 1
            ELSE 0
        END
    ) AS HasSignedNote,
    COUNT(pl.ProcNum) AS ProcCount,
    -- Check 2: Subquery checks for a signed ~GRP~ group note for same patient+date
    -- Using subquery avoids false positives from chained LEFT JOINs
    (
        SELECT COUNT(*)
        FROM procedurelog grp_pl
        INNER JOIN procedurecode grp_pc ON grp_pc.CodeNum = grp_pl.CodeNum
            AND grp_pc.ProcCode = '~GRP~'
        INNER JOIN procnote grp_pn ON grp_pn.ProcNum = grp_pl.ProcNum
        WHERE grp_pl.PatNum = a.PatNum
            AND DATE(grp_pl.ProcDate) = DATE(a.AptDateTime)
            AND grp_pl.AptNum = 0
            AND (
                (grp_pn.Signature IS NOT NULL AND grp_pn.Signature != '')
                OR grp_pn.SigIsTopaz = 1
            )
    ) AS HasSignedGroupNote
FROM appointment a
INNER JOIN patient pat ON pat.PatNum = a.PatNum
INNER JOIN procedurelog pl ON pl.AptNum = a.AptNum
LEFT JOIN procnote pn ON pn.ProcNum = pl.ProcNum
WHERE
    a.AptStatus = 2
    AND a.AptDateTime >= %s
    AND a.AptDateTime < %s
GROUP BY
    a.AptNum,
    a.AptDateTime,
    a.ProvNum,
    a.ProvHyg,
    a.IsHygiene,
    PatientName
HAVING
    ProcCount > 0
    AND HasSignedNote = 0
    AND HasSignedGroupNote = 0
ORDER BY
    a.AptDateTime ASC;
"""

QUERY_PROVIDER_EMAIL = """
SELECT
    p.ProvNum,
    CONCAT(p.FName, ' ', p.LName) AS ProviderName,
    p.Abbr,
    ea.SenderAddress AS EmailAddress
FROM provider p
LEFT JOIN emailaddress ea ON ea.EmailAddressNum = p.EmailAddressNum
WHERE
    p.ProvNum = %s
    AND p.IsHidden = 0;
"""

QUERY_APT_PROCEDURES = """
SELECT
    pl.ProcNum,
    pc.ProcCode,
    pc.Descript AS ProcDescription,
    pl.ToothNum,
    pl.Surf
FROM procedurelog pl
LEFT JOIN procedurecode pc ON pc.CodeNum = pl.CodeNum
WHERE pl.AptNum = %s;
"""

# Get upcoming appointments with D4341/D4342 where patient has insurance
# but no perio chart exists
QUERY_MISSING_PERIO_CHARTS = """
SELECT
    a.AptNum,
    a.AptDateTime,
    a.ProvNum,
    a.ProvHyg,
    CONCAT(pat.LName, ', ', pat.FName) AS PatientName,
    GROUP_CONCAT(DISTINCT pc.ProcCode ORDER BY pc.ProcCode SEPARATOR ', ') AS PerioCodes
FROM appointment a
INNER JOIN patient pat ON pat.PatNum = a.PatNum
INNER JOIN procedurelog pl ON pl.AptNum = a.AptNum
INNER JOIN procedurecode pc ON pc.CodeNum = pl.CodeNum
    AND pc.ProcCode IN ('D4341', 'D4342')
-- Include both treatment planned (TP) and completed procedures
-- ProcStatus: 1=TP, 2=Complete
WHERE pl.ProcStatus IN (1, 2)
-- Patient must have insurance
AND EXISTS (
    SELECT 1 FROM patplan pp
    WHERE pp.PatNum = a.PatNum
    AND pp.IsPending = 0
)
-- Only scheduled appointments
AND a.AptStatus = 1
AND a.AptDateTime >= %s
AND a.AptDateTime <= %s
-- No perio exam on file for this patient
AND NOT EXISTS (
    SELECT 1 FROM perioexam pe
    WHERE pe.PatNum = a.PatNum
    AND pe.ExamDate > '0001-01-01'
)
GROUP BY
    a.AptNum,
    a.AptDateTime,
    a.ProvNum,
    a.ProvHyg,
    PatientName
ORDER BY
    a.AptDateTime ASC;
"""

# =============================================================================
# LOGGING SETUP
# =============================================================================

os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def log(msg):
    """Log to file with full unicode, print to console with ascii fallback."""
    logging.info(msg)
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", errors="replace").decode("ascii"))


# =============================================================================
# DATABASE FUNCTIONS
# =============================================================================

def get_db_connection():
    return mysql.connector.connect(**DB_CONFIG)

def get_unsigned_appointments(cursor, start_date, end_date):
    cursor.execute(QUERY_UNSIGNED_APTS, (start_date, end_date))
    return cursor.fetchall()

def get_provider_info(cursor, prov_num):
    cursor.execute(QUERY_PROVIDER_EMAIL, (prov_num,))
    result = cursor.fetchone()
    if result:
        if not result["EmailAddress"]:
            full_name = result["ProviderName"].strip()
            result["EmailAddress"] = PROVIDER_EMAILS.get(full_name, "")
    return result

def get_apt_procedures(cursor, apt_num):
    cursor.execute(QUERY_APT_PROCEDURES, (apt_num,))
    return cursor.fetchall()

def get_missing_perio_charts(cursor, start_date, end_date):
    cursor.execute(QUERY_MISSING_PERIO_CHARTS, (start_date, end_date))
    return cursor.fetchall()

def generate_provider_notes_file(provider_name, apts, apt_procedures_map, today_str):
    """Generate a pre-filled notes txt file for a provider's unsigned appointments."""
    import os
    import glob
    safe_name    = provider_name.replace(" ", "_").replace("/", "_")
    provider_dir = os.path.join(PROVIDER_NOTES_DIR, safe_name)
    os.makedirs(provider_dir, exist_ok=True)

    # Delete files older than 7 days in this provider's folder
    cutoff = datetime.now() - timedelta(days=NOTES_RETENTION_DAYS)
    for old_file in glob.glob(os.path.join(provider_dir, "*.txt")):
        try:
            file_time = datetime.fromtimestamp(os.path.getmtime(old_file))
            if file_time < cutoff:
                os.remove(old_file)
                log(f"  🗑️ Deleted old notes file: {os.path.basename(old_file)}")
        except Exception:
            pass

    date_str = datetime.now().strftime("%Y-%m-%d")
    filepath = os.path.join(provider_dir, f"{safe_name}_unsigned_notes_{date_str}.txt")

    lines = []
    lines.append("=" * 70)
    lines.append(f"  UNSIGNED NOTES — {provider_name.upper()}")
    lines.append(f"  Generated: {today_str}")
    lines.append(f"  Total appointments: {len(apts)}")
    lines.append("=" * 70)
    lines.append("")
    lines.append("INSTRUCTIONS:")
    lines.append("  - Fill in all fields marked with [brackets]")
    lines.append("  - Copy completed notes into Open Dental for each patient")
    lines.append("  - Sign the note in Open Dental after entry")
    lines.append("")

    for i, apt in enumerate(apts, 1):
        apt_date  = apt["AptDateTime"].strftime("%B %d, %Y")
        apt_time  = apt["AptDateTime"].strftime("%I:%M %p")
        patient   = apt["PatientName"]
        procs     = apt_procedures_map.get(apt["AptNum"], [])

        lines.append("-" * 70)
        lines.append(f"  APPOINTMENT {i} of {len(apts)}")
        lines.append(f"  Date    : {apt_date} at {apt_time}")
        lines.append(f"  Patient : {patient}")
        lines.append(f"  Codes   : {', '.join([p['ProcCode'] for p in procs if p['ProcCode']])}")
        lines.append("-" * 70)
        lines.append("")

        from collections import defaultdict

        QUAD_MAP = {
            "UR": "Upper Right quadrant",
            "UL": "Upper Left quadrant",
            "LR": "Lower Right quadrant",
            "LL": "Lower Left quadrant",
        }

        ADULT_RECALL_CODES = {"D0120","D1110","D0274","D0220","D0230","D0330","D1208","D1310","D1330","D0601"}
        COMP_EXAM_CODES    = {"D0150","D1110","D0274","D0220","D0230","D0330","D1208","D1310","D1330","D0601"}
        CHILD_RECALL_CODES = {"D0120","D1120","D0274","D0220","D0230","D0330","D1208","D1351","D1310","D1330","D0601"}
        CHILD_COMP_CODES   = {"D0150","D1120","D0274","D0220","D0230","D0330","D1208","D1351","D1310","D1330","D0601"}
        LIMITED_EXAM_CODES = {"D0140","D0220","D0230"}
        COMPOSITE_CODES    = {"D2391","D2392","D2393","D2331","D2332","D2335"}
        SRP_CODES          = {"D4341","D4342"}
        EXTRACTION_CODES   = {"D7140","D7210","D7953"}

        code_procs = defaultdict(list)
        apt_codes  = set()
        for proc in procs:
            code = proc.get("ProcCode","")
            if code:
                code_procs[code].append(proc)
                apt_codes.add(code)

        def get_tooth_surf(codes):
            parts = []
            for c in codes:
                for p in code_procs.get(c,[]):
                    t = str(p.get("ToothNum","")).strip()
                    s = str(p.get("Surf","")).strip()
                    if t and t not in ("0",""):
                        parts.append(("#{} ({})".format(t,s)) if s and s not in ("0","") and s not in QUAD_MAP else "#{}".format(t))
            return ", ".join(parts)

        def get_quads(codes):
            quads = []
            for c in codes:
                for p in code_procs.get(c,[]):
                    s = str(p.get("Surf","")).strip().upper()
                    if s in QUAD_MAP and QUAD_MAP[s] not in quads:
                        quads.append(QUAD_MAP[s])
                    elif s and s not in QUAD_MAP and s not in quads:
                        quads.append(s)
            return ", ".join(quads) if quads else "[specify quadrant(s)]"

        def get_xrays(present):
            x = []
            if "D0274" in present: x.append("4 bitewings")
            if "D0220" in present: x.append("periapical(s)")
            if "D0230" in present: x.append("additional periapicals")
            if "D0330" in present: x.append("panoramic")
            return ", ".join(x) if x else "[specify]"

        def write_note(label, detail, template):
            lines.append(">>> {}".format(label))
            if detail:
                lines.append("    Details: {}".format(detail))
            lines.append("")
            lines.append(template)
            lines.append("")
            lines.append("")

        def fill_placeholders(t, tooth_str="", quad_str=""):
            if quad_str:
                t = t.replace("[specify quadrant(s)]", quad_str)
            if tooth_str:
                t = t.replace("[tooth # / surface]", tooth_str)
                t = t.replace("[tooth # / surfaces]", tooth_str)
                t = t.replace("tooth #[__]", "tooth {}".format(tooth_str))
                t = t.replace("at site #[__]", "at site {}".format(tooth_str))
                t = t.replace("on tooth #[__]", "on tooth {}".format(tooth_str))
                t = t.replace("#[__]", tooth_str)
            return t

        notes_written = set()

        # GROUP 1 - Adult Recall
        if apt_codes & ADULT_RECALL_CODES and "D1120" not in apt_codes and "D0150" not in apt_codes:
            present   = apt_codes & ADULT_RECALL_CODES
            codes_str = ", ".join(sorted(present))
            xray_str  = get_xrays(present)
            fluor     = "Fluoride varnish applied." if "D1208" in present else ""
            note = "\n".join([
                "PERIODIC ORAL EVALUATION / RECALL",
                "",
                "Pt presents today for periodic oral evaluation and prophylaxis.",
                "Med hx reviewed with no significant changes noted.",
                "",
                "Radiographs taken and reviewed: {}".format(xray_str),
                "Intraoral and extraoral examination completed.",
                "Periodontal evaluation - [WNL / findings]",
                "Radiographic findings - [describe]",
                "Existing restorations evaluated - [stable / concerns noted]",
                "",
                "Prophylaxis completed with piezo and hand instrumentation.",
                "Teeth polished with prophy cup and prophy paste. Interproximal areas flossed.",
                fluor,
                "OHI provided - brushing and flossing techniques reviewed.",
                "",
                "Discussed findings with patient and reviewed treatment plan.",
                "Discussed risks and benefits of treatment vs. no treatment. Patient questions answered.",
                "",
                "Plan: [list treatment]",
                "",
                "OH - [Good / Fair / Poor]",
                "",
                "NV - [next visit / recall interval]",
            ])
            write_note("RECALL NOTE ({})".format(codes_str), "", note)
            notes_written.update(present)

        # GROUP 2 - Adult Comp Exam
        if "D0150" in apt_codes and "D1120" not in apt_codes and "D0150" not in notes_written:
            present   = apt_codes & COMP_EXAM_CODES
            codes_str = ", ".join(sorted(present))
            xray_str  = get_xrays(present)
            fluor     = "Fluoride varnish applied." if "D1208" in present else ""
            prophy    = ("Prophylaxis completed with piezo and hand instrumentation.\n"
                         "Teeth polished with prophy cup and prophy paste. Interproximal areas flossed.\n"
                         + fluor + "\n"
                         "OHI provided - brushing and flossing techniques reviewed.\n") if "D1110" in present else ""
            note = "\n".join([
                "COMPREHENSIVE ORAL EVALUATION / NEW OR ESTABLISHED PATIENT",
                "",
                "Pt presents today for comprehensive oral evaluation.",
                "Med hx reviewed and documented. [Note any significant medical history]",
                "",
                "Radiographs taken and reviewed: {}".format(xray_str),
                "Full intraoral and extraoral examination completed.",
                "Periodontal screening - [WNL / findings]",
                "Soft tissue evaluation - [WNL / findings]",
                "Occlusal evaluation - [WNL / findings]",
                "Radiographic findings reviewed and discussed with patient.",
                "",
                prophy,
                "Discussed diagnosis and treatment options.",
                "Discussed risks and benefits of treatment vs. no treatment. Patient questions answered.",
                "Patient consent obtained for proposed treatment plan.",
                "",
                "Plan: [list treatment]",
                "",
                "OH - [Good / Fair / Poor]",
                "",
                "NV - [next visit]",
            ])
            write_note("COMPREHENSIVE EXAM NOTE ({})".format(codes_str), "", note)
            notes_written.update(present)

        # GROUP 3 - Child Recall
        if apt_codes & CHILD_RECALL_CODES and "D1120" in apt_codes and "D0150" not in apt_codes and "D0120" not in notes_written:
            present   = apt_codes & CHILD_RECALL_CODES
            codes_str = ", ".join(sorted(present))
            xray_str  = get_xrays(present)
            fluor     = "Fluoride varnish applied." if "D1208" in present else ""
            seals     = "Sealants placed on teeth: [specify teeth]." if "D1351" in present else ""
            ohi       = "Nutritional counseling provided." if "D1310" in present else ""
            note = "\n".join([
                "CHILD PERIODIC EVALUATION / RECALL",
                "",
                "Pt presents today for child periodic evaluation and prophylaxis.",
                "Parent/guardian present. Med hx reviewed with no significant changes noted.",
                "",
                "Radiographs taken and reviewed: {}".format(xray_str),
                "Intraoral and extraoral examination completed.",
                "Periodontal evaluation - [WNL / findings]",
                "Radiographic findings - [describe]",
                "Existing restorations evaluated - [stable / concerns noted]",
                "",
                "Child prophylaxis completed. Teeth polished with prophy cup and prophy paste.",
                "Interproximal areas flossed.",
                fluor,
                seals,
                ohi,
                "OHI provided to patient and parent/guardian - brushing and flossing techniques reviewed.",
                "",
                "Discussed findings with parent/guardian and reviewed treatment plan.",
                "",
                "Plan: [list treatment]",
                "",
                "OH - [Good / Fair / Poor]",
                "",
                "NV - [next visit / recall interval]",
            ])
            write_note("CHILD RECALL NOTE ({})".format(codes_str), "", note)
            notes_written.update(present)

        # GROUP 4 - Child Comp Exam
        if "D0150" in apt_codes and "D1120" in apt_codes and "D0150" not in notes_written:
            present   = apt_codes & CHILD_COMP_CODES
            codes_str = ", ".join(sorted(present))
            xray_str  = get_xrays(present)
            fluor     = "Fluoride varnish applied." if "D1208" in present else ""
            seals     = "Sealants placed on teeth: [specify teeth]." if "D1351" in present else ""
            note = "\n".join([
                "CHILD COMPREHENSIVE ORAL EVALUATION / NEW OR ESTABLISHED PATIENT",
                "",
                "Pt presents today for child comprehensive oral evaluation.",
                "Parent/guardian present. Med hx reviewed and documented.",
                "",
                "Radiographs taken and reviewed: {}".format(xray_str),
                "Full intraoral and extraoral examination completed.",
                "Periodontal screening - [WNL / findings]",
                "Soft tissue evaluation - [WNL / findings]",
                "Occlusal/dental development evaluation - [WNL / findings]",
                "",
                "Child prophylaxis completed. Teeth polished and flossed.",
                fluor,
                seals,
                "OHI provided to patient and parent/guardian.",
                "",
                "Discussed findings and treatment plan with parent/guardian.",
                "Risks, benefits, and alternatives discussed. Questions answered.",
                "",
                "Plan: [list treatment]",
                "",
                "OH - [Good / Fair / Poor]",
                "",
                "NV - [next visit]",
            ])
            write_note("CHILD COMP EXAM NOTE ({})".format(codes_str), "", note)
            notes_written.update(present)

        # GROUP 5 - Limited Exam
        if apt_codes & LIMITED_EXAM_CODES and "D0150" not in apt_codes and "D0120" not in apt_codes and "D0140" not in notes_written:
            present   = apt_codes & LIMITED_EXAM_CODES
            codes_str = ", ".join(sorted(present))
            xray_str  = get_xrays(present)
            note = "\n".join([
                "LIMITED ORAL EVALUATION",
                "",
                "Pt presents today for limited oral evaluation - [chief complaint / tooth #]",
                "Med hx reviewed with no significant changes noted.",
                "",
                "Radiographs taken and reviewed: {}".format(xray_str),
                "Radiographic and intraoral evaluation reveals - [findings]",
                "Recommended tx - [diagnosis and proposed treatment]",
                "",
                "Pt was informed about risks vs. benefits of the proposed treatment,",
                "alternative treatment options, and the option of no treatment.",
                "Patient was informed they may refuse treatment at any point.",
                "",
                "Tx done today - [if any]",
                "",
                "NV - [next visit]",
            ])
            write_note("LIMITED EXAM NOTE ({})".format(codes_str), "", note)
            notes_written.update(present)

        # GROUP 6 - Composites
        comp_present = apt_codes & COMPOSITE_CODES
        if comp_present:
            codes_str  = ", ".join(sorted(comp_present))
            detail_lines = []
            for code in sorted(comp_present):
                if code in code_procs:
                    ct = []
                    for p in code_procs[code]:
                        t = str(p.get("ToothNum","")).strip()
                        s = str(p.get("Surf","")).strip()
                        if t and t not in ("0",""):
                            ct.append("#{} ({})".format(t,s) if s and s not in ("0","") else "#{}".format(t))
                    if ct:
                        detail_lines.append("  {}: {}".format(code, ", ".join(ct)))
            detail_str = "\n".join(detail_lines) if detail_lines else "  {}: [specify teeth]".format(codes_str)
            note = "\n".join([
                "COMPOSITE RESTORATION(S)",
                "",
                "Procedures: {}".format(codes_str),
                "Teeth treated:",
                detail_str,
                "",
                "Med hx reviewed with no significant changes noted.",
                "",
                "Topical anesthesia applied at injection sites.",
                "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.",
                "",
                "Rubber dam / isolation placed.",
                "Caries removed with handpiece and spoon excavator on each tooth.",
                "Cavities prepared and etched. Bonding agent applied and light cured.",
                "Matrix bands and wedges placed as needed for interproximal contacts.",
                "Composite resin placed incrementally and light cured on each restoration.",
                "Restorations contoured, finished, and polished.",
                "Occlusion checked and adjusted as needed on all restored teeth.",
                "",
                "Patient tolerated procedure well.",
                "",
                "NV - [next visit]",
            ])
            tooth_str = get_tooth_surf(comp_present)
            write_note("COMPOSITE NOTE ({})".format(codes_str), tooth_str, note)
            notes_written.update(comp_present)

        # GROUP 7 - SRP
        srp_present = apt_codes & SRP_CODES
        if srp_present:
            codes_str = ", ".join(sorted(srp_present))
            q_str     = get_quads(srp_present)
            note = "\n".join([
                "PERIODONTAL SCALING AND ROOT PLANING",
                "",
                "Procedures: {}".format(codes_str),
                "Quadrant(s): {}".format(q_str),
                "",
                "Med hx reviewed with no significant changes noted.",
                "",
                "Topical anesthesia applied at injection sites.",
                "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.",
                "",
                "SRP completed with piezo ultrasonic and hand instrumentation.",
                "Root surfaces debrided and smoothed.",
                "The area was irrigated with Peridex (chlorhexidine 0.12%).",
                "",
                "Probing depths pre-treatment: [record or reference periodontal chart]",
                "Patient tolerated procedure well.",
                "",
                "Post-op instructions provided:",
                "- Mild soreness expected for 1-2 days",
                "- Rinse with warm salt water 2-3x daily",
                "- Avoid hard/crunchy foods for 24 hours",
                "- Continue prescribed oral hygiene regimen",
                "",
                "NV - [re-evaluation / remaining quads]",
            ])
            write_note("SRP NOTE ({})".format(codes_str), q_str, note)
            notes_written.update(srp_present)

        # GROUP 8 - All Extractions combined
        ext_present = apt_codes & EXTRACTION_CODES
        if ext_present:
            codes_str    = ", ".join(sorted(ext_present))
            tooth_str    = get_tooth_surf(ext_present)
            filled_str   = tooth_str if tooth_str else "[tooth # / teeth]"
            has_surgical = "D7210" in ext_present
            has_graft    = "D7953" in ext_present
            surg_line    = ("Mucoperiosteal flap reflected. Bone removal performed as necessary.\n"
                            "Tooth/teeth sectioned as needed. Elevated and extracted.\n"
                            "Socket(s) debrided and irrigated.") if has_surgical else (
                            "Tooth/teeth elevated and extracted with forceps without complication.\n"
                            "Socket(s) irrigated.")
            graft_line   = ("Ridge preservation bone graft placed at extraction site(s).\n"
                            "Graft material: [specify]. Membrane placed: [yes/no - specify type].") if has_graft else ""
            suture_line  = ("Flap repositioned and sutured with [suture type].\n"
                            "Suture removal scheduled: [date]") if has_surgical else ""
            rx_line      = "Rx: [Amoxicillin 500mg / Ibuprofen 800mg / other]" if has_surgical else "Rx: [if prescribed]"
            ice_line     = "- Ice pack 20 min on/off for first 24 hours" if has_surgical else ""
            note_parts   = [
                "EXTRACTION(S)",
                "",
                "Procedures: {}".format(codes_str),
                "Teeth extracted: {}".format(filled_str),
                "",
                "Med hx reviewed with no significant changes noted.",
                "",
                "Risks, benefits, and alternatives discussed with patient.",
                "Informed consent obtained.",
                "",
                "Topical anesthesia applied at injection sites.",
                "[__] carpules of 2% lidocaine with 1:100,000 epi administered via local infiltration.",
                "",
                surg_line,
            ]
            if graft_line:  note_parts.append(graft_line)
            if suture_line: note_parts.append(suture_line)
            note_parts += [
                "Hemostasis achieved. No complications encountered.",
                "Patient tolerated procedure well.",
                "",
                "Post-op instructions provided (verbal and written):",
                "- Bite on gauze for 30-45 minutes",
                "- Avoid smoking, spitting, straws for 72 hours",
                "- Soft diet recommended",
            ]
            if ice_line: note_parts.append(ice_line)
            note_parts += [
                "- OTC/prescribed pain management as directed",
                "",
                rx_line,
                "",
                "NV - [post-op / as needed]",
            ]
            note = "\n".join(note_parts)
            write_note("EXTRACTION NOTE ({})".format(codes_str), tooth_str, note)
            notes_written.update(ext_present)

        # STANDALONE - remaining codes
        for code in sorted(apt_codes - notes_written):
            if code not in code_procs:
                continue
            proc_list = code_procs[code]
            desc      = proc_list[0].get("ProcDescription","")
            tooth_str = get_tooth_surf({code})
            template  = NOTE_TEMPLATES.get(code, DEFAULT_NOTE_TEMPLATE)
            template  = fill_placeholders(template, tooth_str)
            write_note("NOTE FOR: {} - {}".format(code, desc), tooth_str, template)
            notes_written.add(code)

    lines.append("=" * 70)
    lines.append("  END OF UNSIGNED NOTES")
    lines.append("=" * 70)

    file_content = "\n".join(lines)
    with open(filepath, "w", encoding="utf-8-sig") as f:
        f.write(file_content)

    return filepath

# =============================================================================
# EMAIL BUILDER
# =============================================================================

def build_email_html(provider_name, unsigned_apts, apt_procedures_map, today_str, perio_apts=None):
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)

    todays_apts = []
    carryover_apts = []

    for apt in unsigned_apts:
        apt_date = apt["AptDateTime"].date()
        if apt_date >= yesterday:
            todays_apts.append(apt)
        else:
            carryover_apts.append(apt)

    def apt_rows(apts):
        rows = ""
        for apt in apts:
            apt_date = apt["AptDateTime"].strftime("%b %d, %Y")
            apt_time = apt["AptDateTime"].strftime("%I:%M %p")
            procs = apt_procedures_map.get(apt["AptNum"], [])
            proc_list = ", ".join(
                [f"{p['ProcCode']}" + (f" ({p['ProcDescription']})" if p['ProcDescription'] else "")
                 for p in procs]
            ) or "—"
            rows += f"""
            <tr>
                <td style="padding:8px 12px; border-bottom:1px solid #eee;">{apt_date}</td>
                <td style="padding:8px 12px; border-bottom:1px solid #eee;">{apt_time}</td>
                <td style="padding:8px 12px; border-bottom:1px solid #eee;">{apt['PatientName']}</td>
                <td style="padding:8px 12px; border-bottom:1px solid #eee; color:#666; font-size:13px;">{proc_list}</td>
            </tr>"""
        return rows

    carryover_section = ""
    if carryover_apts:
        carryover_section = f"""
        <h3 style="color:#c0392b; margin-top:24px;">
            [WARN] Carry-Over ({len(carryover_apts)} appointment{'s' if len(carryover_apts)>1 else ''})
        </h3>
        <p style="color:#666; font-size:13px;">These appointments are from previous days and still have no signed note.</p>
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse; margin-top:8px;">
            <thead>
                <tr style="background:#fdf2f2;">
                    <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Date</th>
                    <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Time</th>
                    <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Patient</th>
                    <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Procedures</th>
                </tr>
            </thead>
            <tbody>
                {apt_rows(carryover_apts)}
            </tbody>
        </table>"""

    todays_section = ""
    if todays_apts:
        todays_section = f"""
        <h3 style="color:#2c3e50; margin-top:24px;">
            [NOTES] Recent ({len(todays_apts)} appointment{'s' if len(todays_apts)>1 else ''})
        </h3>
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse; margin-top:8px;">
            <thead>
                <tr style="background:#f0f4f8;">
                    <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Date</th>
                    <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Time</th>
                    <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Patient</th>
                    <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Procedures</th>
                </tr>
            </thead>
            <tbody>
                {apt_rows(todays_apts)}
            </tbody>
        </table>"""

    no_outstanding = ""
    if not carryover_apts and not todays_apts:
        no_outstanding = """
        <div style="text-align:center; padding:40px; color:#27ae60;">
            <div style="font-size:48px;">[OK]</div>
            <p style="font-size:18px; font-weight:bold;">All caught up!</p>
            <p style="color:#666;">No unsigned notes found in the last 30 days.</p>
        </div>"""


    # Build perio chart alert section
    perio_section = ""
    if perio_apts:
        perio_rows = ""
        for apt in perio_apts:
            apt_date = apt["AptDateTime"].strftime("%b %d, %Y")
            apt_time = apt["AptDateTime"].strftime("%I:%M %p")
            perio_rows += f"""
            <tr>
                <td style="padding:8px 12px; border-bottom:1px solid #eee;">{apt_date}</td>
                <td style="padding:8px 12px; border-bottom:1px solid #eee;">{apt_time}</td>
                <td style="padding:8px 12px; border-bottom:1px solid #eee;">{apt["PatientName"]}</td>
                <td style="padding:8px 12px; border-bottom:1px solid #eee; color:#e74c3c; font-size:13px;">{apt["PerioCodes"]}</td>
            </tr>"""

        perio_section = f"""
        <div style="padding:24px 32px; border-top:3px solid #e74c3c;">
            <h3 style="color:#e74c3c; margin-top:0;">
                [PERIO] Missing Perio Charts — Prior Auth Required ({len(perio_apts)} upcoming)
            </h3>
            <p style="color:#666; font-size:13px;">
                The following patients have D4341/D4342 scheduled in the next 30 days,
                have insurance on file, but <strong>no perio chart exists</strong>.
                Please create a perio chart before the appointment to support prior authorization.
            </p>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse; margin-top:8px;">
                <thead>
                    <tr style="background:#fdf2f2;">
                        <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Date</th>
                        <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Time</th>
                        <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Patient</th>
                        <th style="padding:8px 12px; text-align:left; font-size:13px; color:#666;">Procedures</th>
                    </tr>
                </thead>
                <tbody>
                    {perio_rows}
                </tbody>
            </table>
        </div>"""

    total = len(unsigned_apts)
    summary_color = "#27ae60" if total == 0 else "#e74c3c"
    summary_text = "All notes signed [OK]" if total == 0 else f"{total} appointment{'s' if total > 1 else ''} need{'s' if total == 1 else ''} a signed note"

    html = f"""
    <!DOCTYPE html>
    <html>
    <body style="font-family: Arial, sans-serif; margin:0; padding:0; background:#f5f5f5;">
        <div style="max-width:700px; margin:20px auto; background:white; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.1);">
            <div style="background:#2c3e50; padding:24px 32px;">
                <div style="color:#bdc3c7; font-size:13px;">{PRACTICE_NAME}</div>
                <div style="color:white; font-size:22px; font-weight:bold; margin-top:4px;">
                    Daily Note Signature Reminder
                </div>
                <div style="color:#95a5a6; font-size:13px; margin-top:4px;">{today_str}</div>
            </div>
            <div style="background:{summary_color}; padding:12px 32px; color:white; font-size:15px;">
                Dr. {provider_name} — {summary_text}
            </div>
            <div style="padding:24px 32px;">
                <p style="color:#555; font-size:14px; margin-top:0;">
                    The list below shows completed appointments from the <strong>last 30 days</strong>
                    where no signed procedure note was found. Please log into Open Dental and complete
                    any outstanding notes at your earliest convenience.
                </p>
                <p style="color:#888; font-size:12px; background:#fffbe6; padding:10px 14px; border-left:3px solid #f39c12; border-radius:3px;">
                    <strong>Reminder:</strong> At least one procedure in each appointment must have a signed note.
                    A single note or group note covering the entire appointment is sufficient.
                </p>
                {no_outstanding}
                {carryover_section}
                {todays_section}
            </div>

            <!-- Perio Chart Alert Section -->
            {perio_section}

            <div style="background:#f8f9fa; padding:16px 32px; border-top:1px solid #eee;">
                <p style="color:#aaa; font-size:12px; margin:0;">
                    This is an automated message from {PRACTICE_NAME}. Please do not reply to this email.
                </p>
            </div>
        </div>
    </body>
    </html>"""

    return html

# =============================================================================
# EMAIL SENDER
# =============================================================================

def send_email(to_email, provider_name, html_body, unsigned_count, perio_count=0, attachment_path=None):
    today_str = datetime.now().strftime("%B %d, %Y")
    if unsigned_count == 0 and perio_count == 0:
        subject = f"✅ All Clear — {provider_name} | {today_str}"
    elif unsigned_count == 0 and perio_count > 0:
        subject = f"🦷 {perio_count} Missing Perio Chart{'s' if perio_count > 1 else ''} — {provider_name} | {today_str}"
    elif unsigned_count > 0 and perio_count == 0:
        subject = f"⚠️ {unsigned_count} Unsigned Note{'s' if unsigned_count > 1 else ''} — {provider_name} | {today_str}"
    else:
        subject = f"⚠️ {unsigned_count} Unsigned Note{'s' if unsigned_count > 1 else ''} + 🦷 {perio_count} Missing Perio Chart{'s' if perio_count > 1 else ''} — {provider_name} | {today_str}"

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_CONFIG["sender_email"]
    msg["To"]      = to_email
    msg.attach(MIMEText(html_body, "html"))

    # Attach notes file if provided and has content
    if attachment_path and unsigned_count > 0:
        import os
        from email.mime.base import MIMEBase
        from email import encoders
        if os.path.exists(attachment_path):
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            filename = os.path.basename(attachment_path)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_CONFIG["sender_email"], GMAIL_CONFIG["sender_password"])
        server.sendmail(GMAIL_CONFIG["sender_email"], to_email, msg.as_string())

# =============================================================================
# MAIN
# =============================================================================

def main():
    today        = datetime.now().date()
    today_str    = datetime.now().strftime("%B %d, %Y")
    start        = today - timedelta(days=LOOKBACK_DAYS)
    end          = today + timedelta(days=1)
    perio_start  = today
    perio_end    = today + timedelta(days=PERIO_LOOKAHEAD_DAYS)

    log(f"=== Unsigned Notes Emailer Started — {today_str} ===")

    try:
        conn   = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        log("Database connected successfully.")
    except Exception as e:
        log(f"ERROR: Could not connect to database: {e}")
        return

    try:
        unsigned_apts = get_unsigned_appointments(cursor, start, end)
        log(f"Found {len(unsigned_apts)} unsigned appointments total.")
    except Exception as e:
        log(f"ERROR: Query failed: {e}")
        cursor.close()
        conn.close()
        return

    try:
        all_perio_apts = get_missing_perio_charts(cursor, perio_start, perio_end)
        log(f"Found {len(all_perio_apts)} upcoming appointments missing perio charts.")
    except Exception as e:
        log(f"WARNING: Perio chart query failed: {e}")
        all_perio_apts = []

    provider_apts = {}
    for apt in unsigned_apts:
        for prov_field in ["ProvNum", "ProvHyg"]:
            pnum = apt.get(prov_field)
            if pnum and pnum > 0:
                if pnum not in provider_apts:
                    provider_apts[pnum] = []
                if apt["AptNum"] not in [a["AptNum"] for a in provider_apts[pnum]]:
                    provider_apts[pnum].append(apt)

    # Also include providers who have perio alerts but no unsigned notes
    for apt in all_perio_apts:
        for prov_field in ["ProvNum", "ProvHyg"]:
            pnum = apt.get(prov_field)
            if pnum and pnum > 0 and pnum not in provider_apts:
                provider_apts[pnum] = []  # empty unsigned list but still gets email

    log(f"Sending emails to {len(provider_apts)} provider(s).")

    for prov_num, apts in provider_apts.items():
        prov_info = get_provider_info(cursor, prov_num)

        if not prov_info:
            log(f"  Skipping ProvNum {prov_num} — not found or hidden.")
            continue

        provider_name = prov_info["ProviderName"]
        email_address = prov_info["EmailAddress"]

        if not email_address:
            log(f"  Skipping {provider_name} — no email address on file.")
            continue

        try:
            apt_procedures_map = {}
            for apt in apts:
                procs = get_apt_procedures(cursor, apt["AptNum"])
                apt_procedures_map[apt["AptNum"]] = procs
        except Exception as e:
            log(f"  ❌ Failed to get procedures for {provider_name}: {e}")
            continue

        try:
            # Filter perio apts for this provider only
            provider_perio_apts = [
                apt for apt in all_perio_apts
                if apt.get("ProvNum") == prov_num or apt.get("ProvHyg") == prov_num
            ]
            html_body = build_email_html(provider_name, apts, apt_procedures_map, today_str, provider_perio_apts)
        except Exception as e:
            log(f"  ❌ Failed to build email for {provider_name}: {e}")
            import traceback
            log(traceback.format_exc())
            continue

        # Generate notes txt file if provider has unsigned appointments
        notes_filepath = None
        if apts:
            try:
                notes_filepath = generate_provider_notes_file(
                    provider_name, apts, apt_procedures_map, today_str
                )
                log(f"  📄 Notes file generated: {notes_filepath}")
            except Exception as e:
                log(f"  ⚠️ Could not generate notes file for {provider_name}: {e}")

        try:
            send_email(email_address, provider_name, html_body, len(apts), len(provider_perio_apts), notes_filepath)
            log(f"  ✅ Email sent to {provider_name} ({email_address}) — {len(apts)} unsigned, {len(provider_perio_apts)} missing perio charts.")
        except Exception as e:
            log(f"  ❌ Failed to send email to {provider_name} ({email_address}): {e}")
            import traceback
            log(traceback.format_exc())

    cursor.close()
    conn.close()
    log("=== Done ===\n")


if __name__ == "__main__":
    main()
