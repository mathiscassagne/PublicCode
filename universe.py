# Organisé par région / secteur / type

UNIVERSE = [  

    # ══════════════════════════════════════════════════════  
    # 1. ETFs GLOBAUX (base portefeuille)  
    # ══════════════════════════════════════════════════════  
    {"name":"URTH",   "ticker":"urth.us",   "ccy":"USD", "kind":"ETF", "region":"GLOBAL",  "continent":"Global",   "sector":"Broad Market"},  
    {"name":"ACWI",   "ticker":"acwi.us",   "ccy":"USD", "kind":"ETF", "region":"GLOBAL",  "continent":"Global",   "sector":"Broad Market"},  
    {"name":"VT",     "ticker":"vt.us",     "ccy":"USD", "kind":"ETF", "region":"GLOBAL",  "continent":"Global",   "sector":"Broad Market"},  
    {"name":"IWDA",   "ticker":"iwda.uk",   "ccy":"GBP", "kind":"ETF", "region":"GLOBAL",  "continent":"Global",   "sector":"Broad Market"},  

    # ══════════════════════════════════════════════════════  
    # 2. ETFs EUROPE LARGES  
    # ══════════════════════════════════════════════════════  
    {"name":"MEUD",   "ticker":"meud.uk",   "ccy":"GBP", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Broad Market"},  
    {"name":"EZU",    "ticker":"ezu.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Broad Market"},  
    {"name":"VGK",    "ticker":"vgk.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Broad Market"},  
    {"name":"FEZ",    "ticker":"fez.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Broad Market"},  
    {"name":"IEUR",   "ticker":"ieur.us",   "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Broad Market"},  

    # ══════════════════════════════════════════════════════  
    # 3. ETFs EUROPE PAYS  
    # ══════════════════════════════════════════════════════  
    {"name":"EWG",    "ticker":"ewg.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Allemagne"},  
    {"name":"EWF",    "ticker":"ewf.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"France"},  
    {"name":"EWU",    "ticker":"ewu.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"UK"},  
    {"name":"EWI",    "ticker":"ewi.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Italie"},  
    {"name":"EWP",    "ticker":"ewp.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Espagne"},  
    {"name":"EWN",    "ticker":"ewn.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Pays-Bas"},  
    {"name":"EWS_SE", "ticker":"ews.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Suède"},  
    {"name":"EWD",    "ticker":"ewd.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Suède"},  
    {"name":"EWO",    "ticker":"ewo.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Autriche"},  
    {"name":"EDEN",   "ticker":"eden.us",   "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Danemark"},  
    {"name":"EFNL",   "ticker":"efnl.us",   "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Finlande"},  
    {"name":"ENOR",   "ticker":"enor.us",   "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Norvège"},  
    {"name":"EWL",    "ticker":"ewl.us",    "ccy":"USD", "kind":"ETF", "region":"EU",      "continent":"Europe",   "sector":"Suisse"},  

    # ══════════════════════════════════════════════════════  
    # 4. ETFs ASIE LARGES  
    # ══════════════════════════════════════════════════════  
    {"name":"AEJL",   "ticker":"aejl.uk",   "ccy":"GBP", "kind":"ETF", "region":"ASIA",    "continent":"Asie",     "sector":"Broad Market"},  
    {"name":"VAPX",   "ticker":"vapx.us",   "ccy":"USD", "kind":"ETF", "region":"ASIA",    "continent":"Asie",     "sector":"Broad Market"},  
    {"name":"VPL",    "ticker":"vpl.us",    "ccy":"USD", "kind":"ETF", "region":"ASIA",    "continent":"Asie",     "sector":"Broad Market"},  
    {"name":"AAXJ",   "ticker":"aaxj.us",   "ccy":"USD", "kind":"ETF", "region":"ASIA",    "continent":"Asie",     "sector":"Broad Market"},  

    # ══════════════════════════════════════════════════════  
    # 5. ETFs JAPON  
    # ══════════════════════════════════════════════════════  
    {"name":"EWJ",    "ticker":"ewj.us",    "ccy":"USD", "kind":"ETF", "region":"JAPAN",   "continent":"Asie",     "sector":"Japon"},  
    {"name":"DXJ",    "ticker":"dxj.us",    "ccy":"USD", "kind":"ETF", "region":"JAPAN",   "continent":"Asie",     "sector":"Japon"},  
    {"name":"DBJP",   "ticker":"dbjp.us",   "ccy":"USD", "kind":"ETF", "region":"JAPAN",   "continent":"Asie",     "sector":"Japon"},  

    # ══════════════════════════════════════════════════════  
    # 6. ETFs CHINE  
    # ══════════════════════════════════════════════════════  
    {"name":"MCHI",   "ticker":"mchi.us",   "ccy":"USD", "kind":"ETF", "region":"CHINA",   "continent":"Asie",     "sector":"Chine"},  
    {"name":"FXI",    "ticker":"fxi.us",    "ccy":"USD", "kind":"ETF", "region":"CHINA",   "continent":"Asie",     "sector":"Chine"},  
    {"name":"GXC",    "ticker":"gxc.us",    "ccy":"USD", "kind":"ETF", "region":"CHINA",   "continent":"Asie",     "sector":"Chine"},  
    {"name":"CXSE",   "ticker":"cxse.us",   "ccy":"USD", "kind":"ETF", "region":"CHINA",   "continent":"Asie",     "sector":"Chine"},  

    # ══════════════════════════════════════════════════════  
    # 7. ETFs INDE + CORÉE + TAIWAN + ASIE EM  
    # ══════════════════════════════════════════════════════  
    {"name":"INDA",   "ticker":"inda.us",   "ccy":"USD", "kind":"ETF", "region":"INDIA",   "continent":"Asie",     "sector":"Inde"},  
    {"name":"INDY",   "ticker":"indy.us",   "ccy":"USD", "kind":"ETF", "region":"INDIA",   "continent":"Asie",     "sector":"Inde"},  
    {"name":"EWT",    "ticker":"ewt.us",    "ccy":"USD", "kind":"ETF", "region":"TAIWAN",  "continent":"Asie",     "sector":"Taiwan"},  
    {"name":"EWY",    "ticker":"ewy.us",    "ccy":"USD", "kind":"ETF", "region":"KOREA",   "continent":"Asie",     "sector":"Corée"},  
    {"name":"EWM",    "ticker":"ewm.us",    "ccy":"USD", "kind":"ETF", "region":"ASIA_EM", "continent":"Asie",     "sector":"Malaisie"},  
    {"name":"EPHE",   "ticker":"ephe.us",   "ccy":"USD", "kind":"ETF", "region":"ASIA_EM", "continent":"Asie",     "sector":"Philippines"},  
    {"name":"EIDO",   "ticker":"eido.us",   "ccy":"USD", "kind":"ETF", "region":"ASIA_EM", "continent":"Asie",     "sector":"Indonésie"},  
    {"name":"THD",    "ticker":"thd.us",    "ccy":"USD", "kind":"ETF", "region":"ASIA_EM", "continent":"Asie",     "sector":"Thaïlande"},  
    {"name":"VNM",    "ticker":"vnm.us",    "ccy":"USD", "kind":"ETF", "region":"ASIA_EM", "continent":"Asie",     "sector":"Vietnam"},  

    # ══════════════════════════════════════════════════════  
    # 8. ETFs OBLIGATAIRES EUROPE  
    # ══════════════════════════════════════════════════════  
    {"name":"IBTE",   "ticker":"ibte.uk",   "ccy":"GBP", "kind":"BOND", "region":"EU",     "continent":"Europe",   "sector":"Obligations EU"},  
    {"name":"IEAG",   "ticker":"ieag.uk",   "ccy":"GBP", "kind":"BOND", "region":"EU",     "continent":"Europe",   "sector":"Obligations EU"},  
    {"name":"IEGA",   "ticker":"iega.uk",   "ccy":"GBP", "kind":"BOND", "region":"EU",     "continent":"Europe",   "sector":"Obligations EU"},  

    # ══════════════════════════════════════════════════════  
    # 9. ETFs OBLIGATAIRES US (pour ancrage)  
    # ══════════════════════════════════════════════════════  
    {"name":"TLT",    "ticker":"tlt.us",    "ccy":"USD", "kind":"BOND", "region":"US",     "continent":"Amérique", "sector":"Obligations US"},  
    {"name":"IEF",    "ticker":"ief.us",    "ccy":"USD", "kind":"BOND", "region":"US",     "continent":"Amérique", "sector":"Obligations US"},  
    {"name":"HYG",    "ticker":"hyg.us",    "ccy":"USD", "kind":"BOND", "region":"US",     "continent":"Amérique", "sector":"Obligations US"},  
    {"name":"EMB",    "ticker":"emb.us",    "ccy":"USD", "kind":"BOND", "region":"EM",     "continent":"Global",   "sector":"Obligations EM"},  

    # ══════════════════════════════════════════════════════  
    # 10. OR + COMMODITIES  
    # ══════════════════════════════════════════════════════  
    {"name":"GLD",    "ticker":"gld.us",    "ccy":"USD", "kind":"GOLD",      "region":"OTHER",  "continent":"Global",   "sector":"Or"},  
    {"name":"IAU",    "ticker":"iau.us",    "ccy":"USD", "kind":"GOLD",      "region":"OTHER",  "continent":"Global",   "sector":"Or"},  
    {"name":"SLV",    "ticker":"slv.us",    "ccy":"USD", "kind":"COMMODITY", "region":"OTHER",  "continent":"Global",   "sector":"Métaux Précieux"},  
    {"name":"PDBC",   "ticker":"pdbc.us",   "ccy":"USD", "kind":"COMMODITY", "region":"OTHER",  "continent":"Global",   "sector":"Commodities"},  
    {"name":"GSG",    "ticker":"gsg.us",    "ccy":"USD", "kind":"COMMODITY", "region":"OTHER",  "continent":"Global",   "sector":"Commodities"},  
    {"name":"DBC",    "ticker":"dbc.us",    "ccy":"USD", "kind":"COMMODITY", "region":"OTHER",  "continent":"Global",   "sector":"Commodities"},  
    {"name":"COPX",   "ticker":"copx.us",   "ccy":"USD", "kind":"COMMODITY", "region":"OTHER",  "continent":"Global",   "sector":"Cuivre / Mines"},  

    # ══════════════════════════════════════════════════════  
    # 11. STOCKS EUROPE — LARGE CAPS  
    # ══════════════════════════════════════════════════════  
    {"name":"LVMH",   "ticker":"mc.fr",     "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Luxe"},  
    {"name":"ASML",   "ticker":"asml.us",   "ccy":"USD", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Semi-conducteurs"},  
    {"name":"SAP",    "ticker":"sap.us",    "ccy":"USD", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Tech"},  
    {"name":"NESN",   "ticker":"nesn.us",   "ccy":"USD", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Alimentation"},  
    {"name":"NOVN",   "ticker":"nvs.us",    "ccy":"USD", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Pharma"},  
    {"name":"ROCHE",  "ticker":"rhhby.us",  "ccy":"USD", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Pharma"},  
    {"name":"AZN",    "ticker":"azn.us",    "ccy":"USD", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Pharma"},  
    {"name":"SHELL",  "ticker":"shel.us",   "ccy":"USD", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Energie"},  
    {"name":"BP",     "ticker":"bp.uk",     "ccy":"GBP", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Energie"},  
    {"name":"GLEN",   "ticker":"glen.uk",   "ccy":"GBP", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Matières Premières"},  
    {"name":"RIO",    "ticker":"rio.uk",    "ccy":"GBP", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Matières Premières"},  
    {"name":"BHP",    "ticker":"bhp.uk",    "ccy":"GBP", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Matières Premières"},  
    {"name":"BARC",   "ticker":"barc.uk",   "ccy":"GBP", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Finance"},  
    {"name":"HSBA",   "ticker":"hsba.uk",   "ccy":"GBP", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Finance"},  
    {"name":"DBK",    "ticker":"dbk.de",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Finance"},  
    {"name":"BNP",    "ticker":"bnp.fr",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Finance"},  
    {"name":"SAN_FR", "ticker":"san.fr",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Finance"},  
    {"name":"AXA",    "ticker":"cs.fr",     "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Assurance"},  
    {"name":"ALV",    "ticker":"alv.de",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Assurance"},  
    {"name":"SIE",    "ticker":"sie.de",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Industrie"},  
    {"name":"AIR",    "ticker":"air.fr",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Défense / Aero"},  
    {"name":"BAE",    "ticker":"ba.uk",     "ccy":"GBP", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Défense / Aero"},  
    {"name":"VOW3",   "ticker":"vow3.de",   "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Automobile"},  
    {"name":"BMW",    "ticker":"bmw.de",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Automobile"},  
    {"name":"STLA",   "ticker":"stla.us",   "ccy":"USD", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Automobile"},  
    {"name":"KER",    "ticker":"ker.fr",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Luxe"},  
    {"name":"OR",     "ticker":"or.fr",     "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Luxe"},  
    {"name":"DHL",    "ticker":"dpw.de",    "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Logistique"},  
    {"name":"AD",     "ticker":"ad.nl",     "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Distribution"},  
    {"name":"INGA",   "ticker":"inga.nl",   "ccy":"EUR", "kind":"STOCK", "region":"EU",     "continent":"Europe",   "sector":"Finance"},  

    # ══════════════════════════════════════════════════════  
    # 12. STOCKS ASIE — JAPON  
    # ══════════════════════════════════════════════════════  
    {"name":"SONY",   "ticker":"sony.us",   "ccy":"USD", "kind":"STOCK", "region":"JAPAN",  "continent":"Asie",     "sector":"Tech / Electronique"},  
    {"name":"TM",     "ticker":"tm.us",     "ccy":"USD", "kind":"STOCK", "region":"JAPAN",  "continent":"Asie",     "sector":"Automobile"},  
    {"name":"HMC",    "ticker":"hmc.us",    "ccy":"USD", "kind":"STOCK", "region":"JAPAN",  "continent":"Asie",     "sector":"Automobile"},  
    {"name":"FANUY",  "ticker":"fanuy.us",  "ccy":"USD", "kind":"STOCK", "region":"JAPAN",  "continent":"Asie",     "sector":"Industrie / Robotique"},  
    {"name":"KYOCY",  "ticker":"kyocy.us",  "ccy":"USD", "kind":"STOCK", "region":"JAPAN",  "continent":"Asie",     "sector":"Tech"},  
    {"name":"NTDOY",  "ticker":"ntdoy.us",  "ccy":"USD", "kind":"STOCK", "region":"JAPAN",  "continent":"Asie",     "sector":"Jeux Vidéo"},  

    # ══════════════════════════════════════════════════════  
    # 13. STOCKS ASIE — CHINE / HK  
    # ══════════════════════════════════════════════════════  
    {"name":"BABA",   "ticker":"baba.us",   "ccy":"USD", "kind":"STOCK", "region":"CHINA",  "continent":"Asie",     "sector":"Tech / E-commerce"},  
    {"name":"TCEHY",  "ticker":"tcehy.us",  "ccy":"USD", "kind":"STOCK", "region":"CHINA",  "continent":"Asie",     "sector":"Tech"},  
    {"name":"JD",     "ticker":"jd.us",     "ccy":"USD", "kind":"STOCK", "region":"CHINA",  "continent":"Asie",     "sector":"Tech / E-commerce"},  
    {"name":"BIDU",   "ticker":"bidu.us",   "ccy":"USD", "kind":"STOCK", "region":"CHINA",  "continent":"Asie",     "sector":"Tech"},  
    {"name":"NIO",    "ticker":"nio.us",    "ccy":"USD", "kind":"STOCK", "region":"CHINA",  "continent":"Asie",     "sector":"Automobile Electrique"},  

    # ══════════════════════════════════════════════════════  
    # 14. STOCKS ASIE — INDE / CORÉE / TAIWAN  
    # ══════════════════════════════════════════════════════  
    {"name":"INFY",   "ticker":"infy.us",   "ccy":"USD", "kind":"STOCK", "region":"INDIA",  "continent":"Asie",     "sector":"Tech / IT"},  
    {"name":"WIT",    "ticker":"wit.us",    "ccy":"USD", "kind":"STOCK", "region":"INDIA",  "continent":"Asie",     "sector":"Tech / IT"},  
    {"name":"HDB",    "ticker":"hdb.us",    "ccy":"USD", "kind":"STOCK", "region":"INDIA",  "continent":"Asie",     "sector":"Finance"},  
    {"name":"TSM",    "ticker":"tsm.us",    "ccy":"USD", "kind":"STOCK", "region":"TAIWAN", "continent":"Asie",     "sector":"Semi-conducteurs"},  
    {"name":"SAM_KR", "ticker":"ssnlf.us",  "ccy":"USD", "kind":"STOCK", "region":"KOREA",  "continent":"Asie",     "sector":"Tech / Semi-conducteurs"},  

    # ══════════════════════════════════════════════════════  
    # 15. US — MINIMUM (ancrage / benchmark uniquement)  
    # ══════════════════════════════════════════════════════  
    {"name":"SPY",    "ticker":"spy.us",    "ccy":"USD", "kind":"ETF",   "region":"US",     "continent":"Amérique", "sector":"Broad Market"},  
    {"name":"QQQ",    "ticker":"qqq.us",    "ccy":"USD", "kind":"ETF",   "region":"US",     "continent":"Amérique", "sector":"Tech"},  
    {"name":"XOM",    "ticker":"xom.us",    "ccy":"USD", "kind":"STOCK", "region":"US",     "continent":"Amérique", "sector":"Energie"},  
    {"name":"JPM",    "ticker":"jpm.us",    "ccy":"USD", "kind":"STOCK", "region":"US",     "continent":"Amérique", "sector":"Finance"},  
    {"name":"RTX",    "ticker":"rtx.us",    "ccy":"USD", "kind":"STOCK", "region":"US",     "continent":"Amérique", "sector":"Défense / Aero"},  

    # ══════════════════════════════════════════════════════  
    # 16. ETFs SECTORIELS EUROPE  
    # ══════════════════════════════════════════════════════  
    {"name":"EXV1",   "ticker":"exv1.de",   "ccy":"EUR", "kind":"ETF",   "region":"EU",     "continent":"Europe",   "sector":"Banques EU"},  
    {"name":"EXV6",   "ticker":"exv6.de",   "ccy":"EUR", "kind":"ETF",   "region":"EU",     "continent":"Europe",   "sector":"Santé EU"},  
    {"name":"EXH1",   "ticker":"exh1.de",   "ccy":"EUR", "kind":"ETF",   "region":"EU",     "continent":"Europe",   "sector":"Industrie EU"},  
    {"name":"EXV5",   "ticker":"exv5.de",   "ccy":"EUR", "kind":"ETF",   "region":"EU",     "continent":"Europe",   "sector":"Tech EU"},  

    # ══════════════════════════════════════════════════════  
    # 17. CASH  
    # ══════════════════════════════════════════════════════  
    {"name":"CASH",   "ticker": None,       "ccy":"EUR", "kind":"CASH",  "region":"OTHER",  "continent":"Global",   "sector":"Cash"},  
]  

def get_universe(  
    kinds=None,  
    regions=None,  
    continents=None,  
    sectors=None,  
    exclude_names=None,  
    always_include_cash=True,  
):  
    """  
    Filtre l'univers selon des critères.  
    Retourne toujours CASH si always_include_cash=True.  
    """  
    result = []  
    for a in UNIVERSE:  
        if kinds      and a.get("kind")      not in kinds:      continue  
        if regions    and a.get("region")    not in regions:    continue  
        if continents and a.get("continent") not in continents: continue  
        if sectors    and a.get("sector")    not in sectors:    continue  
        if exclude_names and a.get("name")   in exclude_names:  continue  
        result.append(a)  

    if always_include_cash:  
        names = [a["name"] for a in result]  
        if "CASH" not in names:  
            cash = next((a for a in UNIVERSE if a["name"] == "CASH"), None)  
            if cash:  
                result.append(cash)  

    return result 
