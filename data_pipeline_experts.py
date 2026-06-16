#!/usr/bin/env python3
"""
Synthetic EXPERT survey data for:
"Inteligjenca Artificiale ne Sektorin Bankar Shqiptar: Perspektiva e Eksperteve".

This generates a small, hand-curated panel of 10 synthetic expert respondents
(N is intentionally tiny and deliberately composed, unlike the 400-record
customer simulation). It is for disclosed simulation / methodology support only
and is NOT collected data.

Design notes
------------
* Schema matches the real Microsoft Forms Excel/CSV export exactly so the rows
  can sit beside genuine expert responses in analysis.
* Hypothesis H1c is baked in: experts see HIGH potential (Q8) but rate barriers
  as significant (Q11), view the regulatory framework as weak (Q14) and the
  sector as poorly prepared for the EU AI Act (Q15).
* Free-text register is scaled by role: senior execs / the regulator / external
  consultants answer in full and substantively; operational staff are terse and
  blunt, mirroring how real respondents actually wrote.
* Content is grounded in market research (Intesa Sanpaolo AI Lab / "AIxeleration"
  and responsible-AI stance; BKT Smart digital maturity; no standalone Albanian
  AI law; Law 124/2024 on data protection; Bank of Albania as sole banking
  supervisor; EU AI Act classing credit scoring as high-risk from Aug 2026).
"""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any


# --- Exact export schema (column headers from the real Forms export) ---------
H_Q9 = (
    "Renditni keto fusha sipas potencialit qe mendoni se IA ka per sektorin "
    "bankar shqiptar\n(nga me e rendesishmja ne me pak te rendesishmen)"
)
H_BARRIER = "Si i vleresoni pengesat e meposhtme per adoptimin e IA ne sektorin bankar shqiptar?"

EXPORT_HEADERS = [
    "Id",
    "Start time",
    "Completion time",
    "Email",
    "Name",
    "Duke klikuar me poshte, konfirmoj qe kam lexuar informacionin me siper dhe jap pelqimin tim per pjesemarrje ne kete studim.",
    "Cili pershkrim i pershtatet me mire rolit tuaj aktual?",
    "Sa vite pervoje keni ne sektorin bankar apo financiar?",
    "A ka institucioni juaj nje departament ose njesi te dedikuar per transformimin digjital ose inovacionin?",
    "A perdor aktualisht institucioni juaj ndonje teknologji te bazuar ne inteligjence artificiale?",
    "Ne cilat fusha perdoret ose planifikohet perdorimi i IA?",
    "Mund te na tregoni shkurt per nje shembull konkret te perdorimit (ose te planifikimit) te IA ne institucionin tuaj? Cfare ka funksionuar mire dhe cfare jo?",
    "Ne pergjithesi, si e vleresoni potencialin e IA per te transformuar sektorin bankar shqiptar ne 5-10 vitet e ardhshme?",
    H_Q9,
    "A keni vene re ndonje ndryshim ne pritshmeri ose kerkesa te klienteve tuaj lidhur me sherbimet digjitale ose IA ne vitet e fundit?",
    f"{H_BARRIER}.Mungesa e buxhetit te dedikuar per investime ne IA",
    f"{H_BARRIER}.Mungesa e stafit te kualifikuar ne fushen e IA",
    f"{H_BARRIER}.Infrastruktura teknologjike e pamjaftueshme",
    f"{H_BARRIER}.Mungesa e te dhenave cilesore per te trajnuar modele IA",
    f"{H_BARRIER}.Paqartesia e kuadrit rregullator per IA",
    f"{H_BARRIER}.Rezistenca ndaj ndryshimit brenda organizates",
    "Nga pengesat qe vleresuat me siper, cila mendoni se eshte me urgjentja per tu adresuar dhe pse?",
    "A ka shqetesime ne institucionin tuaj mbi privatesine e te dhenave te klienteve, etiken, ose transparencen kur behet fjale per perdorimin e IA ne vendimmarrje?",
    "Si e pershkruani me mire kuadrin aktual rregullator ne Shqiperi per sa i perket perdorimit te IA ne banka?",
    "Sa i pergatitur mendoni se eshte sektori bankar shqiptar per kerkesat e EU AI Act?",
    "Cfare hapash konkrete mendoni se duhet te ndermarre Banka e Shqiperise, AFSA, ose qeveria per te krijuar nje mjedis me te pershtatshem per adoptimin e IA ne banka?",
    "Si e parashikoni sektorin bankar shqiptar pas 5-10 vitesh ne lidhje me IA? Cfare do te kete ndryshuar me shume krahasuar me sot?",
    "Nese do t'i rekomandonit nje banke ne Shqiperi tre hapa konkrete per te filluar (ose perparuar) adoptimin e IA sot, cilat do te ishin ato?",
]

# Canonical option labels (verbatim from the form).
A_CREDIT = "Vleresimi i riskut te kreditit (credit scoring)"
A_FRAUD = "Zbulimi i mashtrimeve (fraud detection)"
A_CHAT = "Sherbimi ndaj klientit (chatbot, asistent virtual)"
A_DATA = "Analiza e te dhenave dhe raportimi"
A_BACK = "Automatizimi i proceseve te brendshme (back-office)"
A_COMPL = "Perputhshmeria rregullatore (compliance, AML/KYC)"
A_MKT = "Marketingu dhe personalizimi i ofertave"

# Q9 ranking items (note: wording differs slightly from Q6 options).
R_CHAT = "Sherbimi ndaj klientit (chatbot, asistente virtuale)"
R_CREDIT = "Vleresimi i riskut te kreditit"
R_FRAUD = "Zbulimi i mashtrimeve"
R_BACK = "Automatizimi i proceseve te brendshme"
R_PERS = "Personalizimi i produkteve per klientin"
R_COMPL = "Perputhshmeria rregullatore"

# Barrier scale labels.
NUK, VOGEL, MOD, MADHE, KRIT = (
    "Nuk eshte pengese",
    "Pengese e vogel",
    "Pengese e moderuar",
    "Pengese e madhe",
    "Pengese kritike",
)
BARRIER_ORDER = ["budget", "staff", "infrastructure", "data", "regulatory", "resistance"]


def barriers(budget, staff, infra, data, reg, resist) -> dict[str, str]:
    return {
        "budget": budget, "staff": staff, "infrastructure": infra,
        "data": data, "regulatory": reg, "resistance": resist,
    }


# --- The curated panel of 10 synthetic experts -------------------------------
# Each persona is a fully specified response. Free-text register varies by role.
PERSONAS: list[dict[str, Any]] = [
    {
        "id": 1, "bank": "Intesa Sanpaolo Bank", "register": "full",
        "start": "4/14/2026 9:42", "completion": "4/14/2026 9:54",
        "Q2_role": "Drejtues i larte (CEO, Zv. Drejtor, Anetar Bordi)",
        "Q3_experience": "Mbi 20 vjet",
        "Q4_dept": "Po, departament te dedikuar",
        "Q5_uses_ai": "Po, ne disa fusha tashme",
        "Q6_areas": [A_FRAUD, A_CHAT, A_DATA, A_COMPL, A_CREDIT],
        "Q7_example": (
            "Ne nivel grupi kemi nje laborator te dedikuar per inteligjencen artificiale dhe programe "
            "te brendshme per adoptimin e saj. Lokalisht kemi nisur me zbulimin e mashtrimeve dhe me "
            "asistente virtuale per pyetjet baze te klienteve. Ka funksionuar mire ulja e kohes se "
            "pergjigjes dhe identifikimi i transaksioneve te dyshimta; sfida kryesore mbetet cilesia e "
            "te dhenave lokale dhe integrimi me sistemet ekzistuese (core banking)."
        ),
        "Q8_potential": 5,
        "Q9_ranking": [R_FRAUD, R_CREDIT, R_CHAT, R_COMPL, R_BACK, R_PERS],
        "Q10_customer_change": (
            "Po, ndjeshem. Klientet presin sherbim 24/7 ne aplikacion, hapje llogarie online dhe "
            "pergjigje te menjehershme. Pritshmeria eshte rritur sidomos te brezi i ri."
        ),
        "Q11_barriers": barriers(MADHE, MADHE, MOD, MADHE, MADHE, MOD),
        "Q12_urgent_barrier": (
            "Paqartesia e kuadrit rregullator se bashku me mungesen e stafit te specializuar. Pa nje "
            "kuader te qarte, edhe kur buxheti ekziston, investimet frenohen nga rreziku ligjor."
        ),
        "Q13_privacy": (
            "Po, privatesia dhe transparenca jane prioritet. Zbatojme Ligjin 124/2024 per mbrojtjen e "
            "te dhenave dhe parimet e grupit per IA te pergjegjshme, me mbikeqyrje njerezore mbi cdo "
            "vendim automatik qe prek klientin."
        ),
        "Q14_regulation": "Ekziston por eshte i pamjaftueshem",
        "Q15_eu_ai_act": 2,
        "Q16_recommendations": (
            "Banka e Shqiperise duhet te nxjerre udhezime specifike per menaxhimin e rrezikut te "
            "modeleve IA, te harmonizoje kuadrin me EU AI Act (sidomos credit scoring si risk i larte) "
            "dhe te krijoje nje 'sandbox' rregullator ku bankat te testojne raste perdorimi te kontrolluara."
        ),
        "Q17_vision": (
            "Pas 5-10 vitesh shumica e nderveprimeve do jene digjitale; vleresimi i rrezikut dhe "
            "kredidhenia do mbeshteten gjeresisht nga IA, ndersa degat fizike do reduktohen dhe do "
            "fokusohen te keshillimi me vlere te shtuar."
        ),
        "Q18_three_steps": (
            "1. Strategji e qarte per IA dhe qeverisje e te dhenave.\n"
            "2. Nis me raste me impakt te larte dhe risk te ulet (chatbot, automatizim back-office).\n"
            "3. Investo ne staf dhe trajnim, ose partneritet me fintech vendase."
        ),
    },
    {
        "id": 2, "bank": "Intesa Sanpaolo Bank", "register": "full",
        "start": "4/16/2026 14:08", "completion": "4/16/2026 14:21",
        "Q2_role": "Drejtues i larte (CEO, Zv. Drejtor, Anetar Bordi)",
        "Q3_experience": "11-20 vjet",
        "Q4_dept": "Po, departament te dedikuar",
        "Q5_uses_ai": "Po, ne disa fusha tashme",
        "Q6_areas": [A_CHAT, A_FRAUD, A_DATA, A_MKT, A_BACK],
        "Q7_example": (
            "Kemi futur asistente virtuale ne kanalet digjitale dhe modele per personalizimin e "
            "ofertave. Asistenti virtual ka ulur ngarkesen ne qendren e thirrjeve, por ende eshte i "
            "kufizuar per kerkesa komplekse dhe kalon te operatori njerezor. Personalizimi kerkon te "
            "dhena me cilesi me te larte se sa kemi sot."
        ),
        "Q8_potential": 5,
        "Q9_ranking": [R_CHAT, R_FRAUD, R_PERS, R_CREDIT, R_BACK, R_COMPL],
        "Q10_customer_change": (
            "Po. Klientet, sidomos ata me te rinj, e konsiderojne aplikacionin si kanalin kryesor dhe "
            "krahasojne sherbimin tone me platformat fintech nderkombetare."
        ),
        "Q11_barriers": barriers(MADHE, MADHE, MOD, MADHE, KRIT, MOD),
        "Q12_urgent_barrier": (
            "Kuadri rregullator. Per nje grup nderkombetar, mungesa e rregullave te qarta lokale per IA "
            "krijon pasiguri ligjore dhe ngadaleson vendimet e investimit."
        ),
        "Q13_privacy": (
            "Po. Cdo perdorim i IA ne vendimmarrje kalon nga nje vleresim i ndikimit mbi privatesine "
            "dhe ruan mbikeqyrjen njerezore. Transparenca ndaj klientit eshte e detyrueshme per ne."
        ),
        "Q14_regulation": "Ekziston por eshte i pamjaftueshem",
        "Q15_eu_ai_act": 2,
        "Q16_recommendations": (
            "Te percaktohet qarte autoriteti pergjegjes per mbikeqyrjen e IA ne banka, te publikohen "
            "udhezime per modelet me risk te larte dhe te nxitet bashkepunimi me rregullatoret e BE-se "
            "perpara afatit te gushtit 2026 te EU AI Act."
        ),
        "Q17_vision": (
            "Banking-u do jete kryesisht 'digital-first', me hapje llogarish dhe kredi te shpejta "
            "online; diferencimi mes bankave do varet nga sa mire perdorin te dhenat dhe IA-ne."
        ),
        "Q18_three_steps": (
            "1. Ndertimi i nje platforme te unifikuar te dhenash.\n"
            "2. Pilotim i kontrolluar ne nje fushe (p.sh. chatbot ose AML).\n"
            "3. Ngritja e qeverisjes se modeleve dhe trajnimi i stafit."
        ),
    },
    {
        "id": 3, "bank": "Intesa Sanpaolo Bank", "register": "full",
        "start": "4/22/2026 11:30", "completion": "4/22/2026 11:41",
        "Q2_role": "Menaxher IT / Teknologjie / Inovacioni",
        "Q3_experience": "11-20 vjet",
        "Q4_dept": "Po, departament te dedikuar",
        "Q5_uses_ai": "Po, ne disa fusha tashme",
        "Q6_areas": [A_FRAUD, A_DATA, A_BACK, A_COMPL],
        "Q7_example": (
            "Perdorim modele per zbulimin e mashtrimeve dhe automatizimin e raportimeve rregullatore. "
            "Rezultatet ne fraud detection kane qene premtuese, por modelet kerkojne ritrajnim te "
            "vazhdueshem dhe nje infrastrukture te dhenash qe ende po e konsolidojme. Integrimi me "
            "sistemet legacy mbetet pjesa me e veshtire."
        ),
        "Q8_potential": 4,
        "Q9_ranking": [R_FRAUD, R_BACK, R_CREDIT, R_COMPL, R_CHAT, R_PERS],
        "Q10_customer_change": (
            "Po, kerkesa per sherbime te menjehershme dhe vetesherbim ne aplikacion eshte rritur; "
            "klientet tolerojne gjithnje e me pak nderprerjet."
        ),
        "Q11_barriers": barriers(MOD, KRIT, KRIT, MADHE, MADHE, MOD),
        "Q12_urgent_barrier": (
            "Mungesa e stafit te kualifikuar ne IA dhe infrastruktura. Pa inxhiniere te dhenash dhe "
            "platforme te qendrueshme, projektet mbeten ne faze pilot dhe nuk shkojne ne prodhim."
        ),
        "Q13_privacy": (
            "Po. Aplikojme minimizimin e te dhenave dhe enkriptim perpara se te dhenat te perpunohen nga "
            "modelet; vendimet me ndikim te larte mbahen nen mbikeqyrje njerezore."
        ),
        "Q14_regulation": "Ekziston por eshte i pamjaftueshem",
        "Q15_eu_ai_act": 3,
        "Q16_recommendations": (
            "Standardizimi i qeverisjes se modeleve IA, kerkesa minimale per dokumentim teknik dhe "
            "log-im te mbikeqyrjes njerezore, si dhe nxitja e ndarjes se anonimizuar te te dhenave mes "
            "institucioneve per trajnim."
        ),
        "Q17_vision": (
            "Shumica e proceseve te brendshme do jene te automatizuara dhe vendimet do mbeshteten nga "
            "modele; do rritet roli i ekipeve te te dhenave brenda bankave."
        ),
        "Q18_three_steps": (
            "1. Konsolidim i te dhenave dhe nje platforme e vetme.\n"
            "2. Punesim/trajnim i inxhiniereve te te dhenave.\n"
            "3. Nje rast perdorimi ne prodhim me matje te qarte impakti."
        ),
    },
    {
        "id": 4, "bank": "BKT (Banka Kombetare Tregtare)", "register": "full",
        "start": "4/28/2026 16:12", "completion": "4/28/2026 16:23",
        "Q2_role": "Menaxher IT / Teknologjie / Inovacioni",
        "Q3_experience": "6-10 vjet",
        "Q4_dept": "Po, por eshte pjese e departamentit te IT",
        "Q5_uses_ai": "Jo, por kemi plane konkrete per vitin e ardhshem",
        "Q6_areas": [A_CHAT, A_BACK],
        "Q7_example": (
            "Aktualisht jemi te forte ne banking-un digjital me aplikacionin tone mobile, por IA-ne "
            "ende nuk e perdorim ne prodhim. Per vitin e ardhshem planifikojme nje chatbot per pyetjet "
            "baze dhe automatizim te disa proceseve back-office. Sfida me e madhe eshte nderlidhja me "
            "core banking system-in dhe sigurimi i te dhenave te paster."
        ),
        "Q8_potential": 4,
        "Q9_ranking": [R_CHAT, R_BACK, R_FRAUD, R_CREDIT, R_PERS, R_COMPL],
        "Q10_customer_change": (
            "Po. Klientet perdorin mobile banking-un 24/7 dhe kerkojne hapje llogarie e pagesa pa "
            "ardhur ne dege. Numri i vizitave ne sportel ka rene."
        ),
        "Q11_barriers": barriers(MADHE, MADHE, KRIT, MADHE, MOD, MADHE),
        "Q12_urgent_barrier": (
            "Infrastruktura teknologjike. Pa nje platforme te dhenash dhe integrim te mire me sistemet "
            "ekzistuese, cdo projekt IA ngec ne faze testimi."
        ),
        "Q13_privacy": (
            "Po, ka shqetesime per privatesine. Mendojme t'i adresojme me politika te brendshme, "
            "anonimizim te te dhenave per trajnim dhe akses te kufizuar."
        ),
        "Q14_regulation": "Ekziston por eshte i pamjaftueshem",
        "Q15_eu_ai_act": 2,
        "Q16_recommendations": (
            "Krijimi i nje kuadri rregullator te qarte per IA, fillimisht ne nje mjedis testimi, dhe "
            "me pas futja graduale ne sektore te caktuar; gjithashtu mbeshtetje per ngritjen e kapaciteteve."
        ),
        "Q17_vision": (
            "IA do perdoret me shume per sherbime specifike si chatbot me klientet dhe automatizim "
            "sherbimesh; degat do shnderrohen ne pika keshillimi."
        ),
        "Q18_three_steps": (
            "1. Ngritja e infrastruktures se nevojshme.\n"
            "2. Marrja/trajnimi i stafit perkates.\n"
            "3. Testimi i nje sherbimi te caktuar me IA perpara prodhimit."
        ),
    },
    {
        "id": 5, "bank": "BKT (Banka Kombetare Tregtare)", "register": "full",
        "start": "5/4/2026 10:05", "completion": "5/4/2026 10:18",
        "Q2_role": "Drejtues i larte (CEO, Zv. Drejtor, Anetar Bordi)",
        "Q3_experience": "11-20 vjet",
        "Q4_dept": "Po, por eshte pjese e departamentit te IT",
        "Q5_uses_ai": "Jo, por e kemi ne diskutim",
        "Q6_areas": [A_FRAUD, A_CREDIT, A_BACK],
        "Q7_example": (
            "Si banka me e madhe ne treg kemi mundesi te mira investimi, por po veprojme me kujdes. "
            "Po diskutojme perdorimin e IA per vleresimin e rrezikut te kreditit dhe zbulimin e "
            "mashtrimeve. Ende nuk kemi nje rast ne prodhim sepse duam fillimisht nje strategji te "
            "qarte dhe nje kuader te brendshem qeverisjeje."
        ),
        "Q8_potential": 5,
        "Q9_ranking": [R_CREDIT, R_FRAUD, R_BACK, R_CHAT, R_COMPL, R_PERS],
        "Q10_customer_change": (
            "Po, klientet kerkojne procese me te shpejta kredie dhe pagesash dhe me pak burokraci ne dege."
        ),
        "Q11_barriers": barriers(MOD, MADHE, MADHE, MADHE, MADHE, MADHE),
        "Q12_urgent_barrier": (
            "Stafi i kualifikuar dhe paqartesia rregullatore. Buxhetin e kemi, por mungojne specialistet "
            "dhe nje kuader ligjor qe te na japi siguri per kredidhenien me IA."
        ),
        "Q13_privacy": (
            "Po, ka shqetesime te konsiderueshme, sidomos per vendimet e kredise. Mendojme t'i adresojme "
            "me transparence ndaj klientit dhe me te drejten per shqyrtim njerezor te vendimit."
        ),
        "Q14_regulation": "Praktikisht nuk ekziston specifikisht per IA",
        "Q15_eu_ai_act": 2,
        "Q16_recommendations": (
            "Hartimi i nje kuadri rregullator specifik per IA ne banka, percaktimi i pergjegjesive mes "
            "Bankes se Shqiperise dhe AFSA-s, dhe nxitja e investimeve ne kapacitete njerezore."
        ),
        "Q17_vision": (
            "Kreditimi do behet me inteligjent dhe me i shpejte; konkurrenca do zhvendoset nga rrjeti "
            "i degeve te aftesia per te perdorur te dhenat."
        ),
        "Q18_three_steps": (
            "1. Strategji e qarte IA dhe infrastrukture te dhenash.\n"
            "2. Fillim me raste me impakt te larte dhe risk te ulet.\n"
            "3. Ndertim i kapaciteteve te brendshme ose partneritet."
        ),
    },
    {
        "id": 6, "bank": "Banka e Shqiperise / AFSA", "register": "full",
        "start": "5/9/2026 13:20", "completion": "5/9/2026 13:38",
        "Q2_role": "Perfaqesues rregullator (Banka e Shqiperise, AFSA, ose institucion tjeter mbikeqyres)",
        "Q3_experience": "11-20 vjet",
        "Q4_dept": "Po, por eshte pjese e departamentit te IT",
        "Q5_uses_ai": "Jo, por e kemi ne diskutim",
        "Q6_areas": [A_COMPL, A_FRAUD, A_DATA],
        "Q7_example": (
            "Nga kendveshtrimi mbikeqyres, interesi kryesor eshte perdorimi i IA per perputhshmerine "
            "rregullatore (AML/KYC) dhe monitorimin e transaksioneve. Vleresojme se teknologjia ndihmon, "
            "por mungon ende nje kuader i qarte per mbikeqyrjen e modeleve dhe per pergjegjesine ne rast "
            "vendimesh te gabuara."
        ),
        "Q8_potential": 4,
        "Q9_ranking": [R_COMPL, R_FRAUD, R_CREDIT, R_BACK, R_CHAT, R_PERS],
        "Q10_customer_change": (
            "Po. Verejme presion ne rritje nga konsumatoret per sherbime digjitale, gje qe rrit edhe "
            "rendesine e mbrojtjes se te dhenave dhe transparences ne vendimmarrje."
        ),
        "Q11_barriers": barriers(MOD, MADHE, MOD, MADHE, KRIT, MOD),
        "Q12_urgent_barrier": (
            "Paqartesia e kuadrit rregullator. Pa rregulla te qarta per IA - sidomos per sistemet me "
            "risk te larte si credit scoring - bankat veprojne me kujdes te tepruar dhe mbikeqyrja "
            "behet e veshtire. Kjo eshte pengesa qe duhet adresuar e para."
        ),
        "Q13_privacy": (
            "Po, ky eshte nje shqetesim qendror. Ligji 124/2024 per mbrojtjen e te dhenave vendos "
            "kerkesa per vendimmarrjen e automatizuar dhe profilizimin. Adresimi kerkon vleresime te "
            "ndikimit mbi mbrojtjen e te dhenave, transparence ndaj klientit dhe mbikeqyrje njerezore "
            "te detyrueshme per vendimet me ndikim te larte."
        ),
        "Q14_regulation": "Praktikisht nuk ekziston specifikisht per IA",
        "Q15_eu_ai_act": 2,
        "Q16_recommendations": (
            "1) Banka e Shqiperise te publikoje udhezime per menaxhimin e rrezikut te modeleve IA. "
            "2) Harmonizim i kuadrit me EU AI Act, duke klasifikuar credit scoring-un si risk te larte "
            "perpara afatit te gushtit 2026. 3) Krijimi i nje 'regulatory sandbox' dhe rritja e "
            "kapaciteteve mbikeqyrese te institucioneve."
        ),
        "Q17_vision": (
            "Pres nje kuader rregullator me te pjekur, te perafruar me BE-ne, dhe perdorim me te gjere "
            "te IA ne AML dhe credit scoring nen mbikeqyrje me te forte. Sfida do jete balanca mes "
            "inovacionit dhe mbrojtjes se konsumatorit."
        ),
        "Q18_three_steps": (
            "1. Vleresim i rrezikut dhe inventar i sistemeve IA ekzistuese.\n"
            "2. Politika e brendshme per qeverisjen e modeleve dhe te dhenave.\n"
            "3. Fillim me raste me risk te ulet, me dokumentim e mbikeqyrje njerezore."
        ),
    },
    {
        "id": 7, "bank": "Konsulence / fintech", "register": "full",
        "start": "5/12/2026 18:44", "completion": "5/12/2026 18:55",
        "Q2_role": "Konsulent ose ekspert i jashtem (fintech, teknologji, financa)",
        "Q3_experience": "6-10 vjet",
        "Q4_dept": "Jo, por ka plane per te krijuar",
        "Q5_uses_ai": "Po, ne disa fusha tashme",
        "Q6_areas": [A_CHAT, A_BACK, A_DATA, A_MKT],
        "Q7_example": (
            "Kam ndihmuar disa institucione te vendosin asistente virtuale dhe lexues te automatizuar "
            "te formularve. Aty ku te dhenat ishin te paster, rezultati ka qene shume i mire; aty ku "
            "te dhenat ishin te shperndara dhe pa cilesi, projektet kane deshtuar ose jane vonuar. "
            "Mesimi kryesor: pa qeverisje te te dhenave, IA nuk jep vlere."
        ),
        "Q8_potential": 5,
        "Q9_ranking": [R_CHAT, R_FRAUD, R_BACK, R_CREDIT, R_PERS, R_COMPL],
        "Q10_customer_change": (
            "Po, qarte. Klientet i krahasojne bankat me aplikacionet fintech dhe presin te njejten "
            "shpejtesi e thjeshtesi."
        ),
        "Q11_barriers": barriers(MADHE, KRIT, MADHE, KRIT, MADHE, MOD),
        "Q12_urgent_barrier": (
            "Cilesia e te dhenave dhe stafi. Shumica e bankave shqiptare nuk kane ende te dhena te "
            "gatshme per trajnim modelesh dhe i mungojne inxhinieret e te dhenave; keto duhen adresuar "
            "para se te flitet per IA te avancuar."
        ),
        "Q13_privacy": (
            "Po. Keshilloj gjithmone enkriptim, minimizim te te dhenave dhe nje 'human-in-the-loop' per "
            "vendimet qe prekin klientin; ndryshe rreziku reputacional dhe ligjor eshte i larte."
        ),
        "Q14_regulation": "Ekziston por eshte i pamjaftueshem",
        "Q15_eu_ai_act": 2,
        "Q16_recommendations": (
            "Udhezime praktike per bankat e vogla, lehtesira fiskale per investime ne teknologji dhe "
            "trajnim, dhe nje sandbox ku fintech-et e bankat te testojne bashke pa rrezik rregullator."
        ),
        "Q17_vision": (
            "Bankat qe investojne sot ne te dhena do dominojne; ato qe presin do mbeten thjesht "
            "perdorues te zgjidhjeve te jashtme. Do shohim me shume partneritete banke-fintech."
        ),
        "Q18_three_steps": (
            "1. Pastrimi dhe qeverisja e te dhenave.\n"
            "2. Nje rast perdorimi i thjeshte me ROI te matshem (p.sh. chatbot).\n"
            "3. Trajnim i stafit ose marrje e ekspertizes nga jashte."
        ),
    },
    {
        "id": 8, "bank": "Konsulence", "register": "medium",
        "start": "5/15/2026 12:02", "completion": "5/15/2026 12:10",
        "Q2_role": "Konsulent ose ekspert i jashtem (fintech, teknologji, financa)",
        "Q3_experience": "3-5 vjet",
        "Q4_dept": "Jo, por ka plane per te krijuar",
        "Q5_uses_ai": "Jo, por kemi plane konkrete per vitin e ardhshem",
        "Q6_areas": [A_CHAT, A_BACK],
        "Q7_example": (
            "Kryesisht kemi pilotuar automatizim te proceseve back-office dhe nje chatbot baze. Ka "
            "ndihmuar ne shpejtesi, por ende kerkon shume mbikeqyrje njerezore."
        ),
        "Q8_potential": 4,
        "Q9_ranking": [R_BACK, R_CHAT, R_FRAUD, R_CREDIT, R_PERS, R_COMPL],
        "Q10_customer_change": "Po, klientet duan procese me te shpejta dhe online.",
        "Q11_barriers": barriers(MADHE, MADHE, MADHE, MADHE, MOD, MOD),
        "Q12_urgent_barrier": (
            "Buxheti dhe stafi. Bankat e vogla e kane te veshtire te financojne dhe te gjejne staf per IA."
        ),
        "Q13_privacy": "Po, ka shqetesime; duhet kujdes me te dhenat personale.",
        "Q14_regulation": "Nuk kam informacion te mjaftueshem per te vleresuar",
        "Q15_eu_ai_act": 3,
        "Q16_recommendations": (
            "Informim me i mire, rregullore te percaktuara qarte dhe mbeshtetje per trajnimin e stafit."
        ),
        "Q17_vision": "Do kete me shume automatizim dhe me pak pune manuale ne sportel.",
        "Q18_three_steps": "1. Trajnim stafi. 2. Nje rast pilot. 3. Investim ne infrastrukture.",
    },
    {
        "id": 9, "bank": "Credins Bank", "register": "terse",
        "start": "5/18/2026 9:15", "completion": "5/18/2026 9:21",
        "Q2_role": "Specialist Bankar",
        "Q3_experience": "6-10 vjet",
        "Q4_dept": "Po, por eshte pjese e departamentit te IT",
        "Q5_uses_ai": "Jo, por e kemi ne diskutim",
        "Q6_areas": [],
        "Q7_example": "Aktualisht nuk perdorim. Eshte ne diskutim per sherbimin ndaj klientit.",
        "Q8_potential": 3,
        "Q9_ranking": [R_CHAT, R_FRAUD, R_BACK, R_CREDIT, R_PERS, R_COMPL],
        "Q10_customer_change": "Po, kerkojne me shume sherbime online.",
        "Q11_barriers": barriers(MADHE, MADHE, MADHE, MOD, MOD, VOGEL),
        "Q12_urgent_barrier": "Buxheti dhe infrastruktura.",
        "Q13_privacy": "Po.",
        "Q14_regulation": "Nuk kam informacion te mjaftueshem per te vleresuar",
        "Q15_eu_ai_act": 2,
        "Q16_recommendations": "Rregullore me te qarta dhe trajnim stafi.",
        "Q17_vision": "Me pak radhe ne sportel, me shume sherbim ne aplikacion.",
        "Q18_three_steps": "Trajnim stafi, instalim sistemesh, nje sherbim pilot.",
    },
    {
        "id": 10, "bank": "Tirana Bank", "register": "terse",
        "start": "5/20/2026 15:48", "completion": "5/20/2026 15:53",
        "Q2_role": "Operacione",
        "Q3_experience": "3-5 vjet",
        "Q4_dept": "Jo, por ka plane per te krijuar",
        "Q5_uses_ai": "Nuk jam i/e sigurt",
        "Q6_areas": [],
        "Q7_example": "Ne shqyrtim. Mendojme per automatizim te proceseve ne sportel.",
        "Q8_potential": 3,
        "Q9_ranking": [R_BACK, R_CHAT, R_FRAUD, R_CREDIT, R_PERS, R_COMPL],
        "Q10_customer_change": "Po pak.",
        "Q11_barriers": barriers(MADHE, MOD, MADHE, MADHE, MOD, MOD),
        "Q12_urgent_barrier": "Mungesa e buxhetit per investime.",
        "Q13_privacy": "Jo deri tani.",
        "Q14_regulation": "Nuk kam informacion te mjaftueshem per te vleresuar",
        "Q15_eu_ai_act": 3,
        "Q16_recommendations": "Informim dhe rregullore te percaktuara.",
        "Q17_vision": "Me pak radhe ne sportel.",
        "Q18_three_steps": "Trajnim stafi, instalim i sistemeve.",
    },
]


def to_record(p: dict[str, Any]) -> dict[str, Any]:
    """Structured record for the review UI / JSON output."""
    return {
        "metadata": {
            "id": p["id"],
            "synthetic": True,
            "purpose": "disclosed synthetic expert respondent for methodology support only",
            "role_persona": p["Q2_role"],
            "bank": p["bank"],
            "register": p["register"],
        },
        "answers": {
            "start_time": p["start"],
            "completion_time": p["completion"],
            "Q1_consent": "Po, jap pelqimin tim",
            "Q2_role": p["Q2_role"],
            "Q3_experience": p["Q3_experience"],
            "Q4_dept": p["Q4_dept"],
            "Q5_uses_ai": p["Q5_uses_ai"],
            "Q6_areas": p["Q6_areas"],
            "Q7_example": p["Q7_example"],
            "Q8_potential": p["Q8_potential"],
            "Q9_ranking": p["Q9_ranking"],
            "Q10_customer_change": p["Q10_customer_change"],
            "Q11_barriers": p["Q11_barriers"],
            "Q12_urgent_barrier": p["Q12_urgent_barrier"],
            "Q13_privacy": p["Q13_privacy"],
            "Q14_regulation": p["Q14_regulation"],
            "Q15_eu_ai_act": p["Q15_eu_ai_act"],
            "Q16_recommendations": p["Q16_recommendations"],
            "Q17_vision": p["Q17_vision"],
            "Q18_three_steps": p["Q18_three_steps"],
        },
    }


def export_row(p: dict[str, Any]) -> dict[str, Any]:
    """Row matching the exact Microsoft Forms export schema."""
    areas = ";".join(p["Q6_areas"]) + (";" if p["Q6_areas"] else "")
    ranking = ";".join(p["Q9_ranking"]) + ";"
    b = p["Q11_barriers"]
    values = [
        p["id"], p["start"], p["completion"], "anonymous", "",
        "Po, jap pelqimin tim",
        p["Q2_role"], p["Q3_experience"], p["Q4_dept"], p["Q5_uses_ai"],
        areas, p["Q7_example"], p["Q8_potential"], ranking, p["Q10_customer_change"],
        b["budget"], b["staff"], b["infrastructure"], b["data"], b["regulatory"], b["resistance"],
        p["Q12_urgent_barrier"], p["Q13_privacy"], p["Q14_regulation"], p["Q15_eu_ai_act"],
        p["Q16_recommendations"], p["Q17_vision"], p["Q18_three_steps"],
    ]
    return dict(zip(EXPORT_HEADERS, values))


def summarize(records: list[dict[str, Any]]) -> dict[str, Any]:
    n = len(records)
    pot = [r["answers"]["Q8_potential"] for r in records]
    eu = [r["answers"]["Q15_eu_ai_act"] for r in records]
    sev = {NUK: 0, VOGEL: 1, MOD: 2, MADHE: 3, KRIT: 4}
    barrier_means: dict[str, float] = {}
    for key in BARRIER_ORDER:
        vals = [sev[r["answers"]["Q11_barriers"][key]] for r in records]
        barrier_means[key] = round(sum(vals) / n, 2)
    role_counts: dict[str, int] = {}
    for r in records:
        role_counts[r["answers"]["Q2_role"]] = role_counts.get(r["answers"]["Q2_role"], 0) + 1
    return {
        "n": n,
        "synthetic_only": True,
        "role_counts": dict(sorted(role_counts.items())),
        "H1c_checks": {
            "potential_mean_Q8": round(sum(pot) / n, 2),
            "eu_ai_act_readiness_mean_Q15": round(sum(eu) / n, 2),
            "barrier_severity_means_0to4": barrier_means,
            "interpretation": (
                "H1c: experts recognise high potential (high Q8) yet rate barriers as "
                "significant and the sector as poorly prepared for the EU AI Act (low Q15)."
            ),
        },
    }


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=EXPORT_HEADERS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    outdir = Path("output")
    records = [to_record(p) for p in PERSONAS]
    rows = [export_row(p) for p in PERSONAS]
    summary = summarize(records)

    write_csv(outdir / "experts_synthetic.csv", rows)
    write_json(outdir / "experts_records.json", records)
    write_json(outdir / "experts_summary.json", summary)

    print(json.dumps({"n": len(records), "summary": summary}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
