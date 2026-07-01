using OutlookShredder.Proxy.Models.Drawing;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Bilingual (English / Spanish) labels for the shop-facing drawing output. The shop floor reads the
/// printed PDF, so every on-drawing LABEL shows both languages as "English / Spanish" via <see cref="T"/>.
/// Out of scope (stay English): the Pixar client UI, the FAB-note grammar (parser input), and the
/// engineering spec-summary readout in the footnote box.
///
/// REVIEW: the Spanish strings below are a first draft — have the shop confirm the trade terms.
/// </summary>
internal static class Bi
{
    // key -> (English, Spanish). T(key) composes "English / Spanish".
    private static readonly Dictionary<string, (string En, string Es)> Map = new()
    {
        // ── view / panel titles ─────────────────────────────────────────────
        ["flatPattern.cut"]      = ("Flat pattern (cut)",        "Patrón plano (corte)"),
        ["formedPart"]           = ("Formed part",               "Pieza formada"),
        ["endSection"]           = ("End section",               "Sección de extremo"),
        ["sideSection"]          = ("Side section",              "Sección lateral"),
        ["plate.topView"]        = ("Plate — top view",          "Placa — vista superior"),
        ["basePlate.topView"]    = ("Base plate",                "Placa base"),
        ["bearingPlate.topView"] = ("Bearing plate",             "Placa de apoyo"),
        ["columnElevation"]      = ("Column elevation",          "Alzado de columna"),
        ["spade.faceView"]       = ("Spade — face view",         "Disco — vista frontal"),
        ["miter.crossSection"]   = ("Cross-section",             "Sección transversal"),
        ["miter.elevation"]      = ("Elevation — outside length + end angles", "Alzado — largo exterior + ángulos de extremo"),
        ["miter.face"]           = ("MITER FACE",                "CARA DE INGLETE"),
        ["miter.isometric"]      = ("Isometric — bevel detail, both ends", "Isométrico — detalle del bisel, ambos extremos"),

        // ── callouts ────────────────────────────────────────────────────────
        ["finish"]               = ("Finish",                    "Acabado"),
        ["polish.direction"]     = ("Polish direction",          "Dirección de pulido"),
        ["finish.outside"]       = ("Finish: outside",           "Acabado: exterior"),
        ["finish.inside"]        = ("Finish: inside",            "Acabado: interior"),
        ["hem"]                  = ("HEM",                       "DOBLADILLO"),
        ["ret"]                  = ("ret",                       "ret"),
        ["tube"]                 = ("tube",                      "tubo"),
        ["tube.cap"]             = ("Tube",                      "Tubo"),
        ["pipe.cap"]             = ("Pipe",                      "Tubería"),

        // ── section-cuts key ────────────────────────────────────────────────
        ["sectionCuts"]          = ("Section cuts",              "Cortes de sección"),
        ["side"]                 = ("Side",                      "Lateral"),
        ["end"]                  = ("End",                       "Extremo"),

        // ── header / blank-size line ────────────────────────────────────────
        ["flatBlank"]            = ("Flat blank",                "Desarrollo plano"),
        ["cutTo"]                = ("Cut to",                    "Cortar a"),
        ["plateToCut"]           = ("Plate to cut",              "Placa a cortar"),
        ["thickness"]            = ("Thickness",                 "Espesor"),
        ["holes"]                = ("holes",                     "agujeros"),
        ["dia"]                  = ("dia",                       "diám"),
        ["toEnd"]                = ("to end",                    "al extremo"),

        // ── spec table (attribute labels) ───────────────────────────────────
        ["spec.quantity"]        = ("Quantity",                  "Cantidad"),
        ["spec.web"]             = ("Web",                       "Alma"),
        ["spec.flanges"]         = ("Flanges",                   "Alas"),
        ["spec.legs"]            = ("Legs",                       "Patas"),
        ["spec.length"]          = ("Length",                    "Largo"),
        ["spec.width"]           = ("Width",                     "Ancho"),
        ["spec.depth"]           = ("Depth",                     "Profundidad"),
        ["spec.material"]        = ("Material",                  "Material"),
        ["spec.od"]              = ("OD",                        "DE"),
        ["spec.nps"]             = ("NPS",                       "NPS"),
        ["spec.class"]           = ("Class",                     "Clase"),
        ["spec.height"]          = ("Height",                    "Altura"),
        ["spec.section"]         = ("Section",                   "Sección"),
        ["spec.wall"]            = ("Wall",                      "Pared"),

        // ── footnote legend tokens ──────────────────────────────────────────
        ["legend.solidCut"]      = ("solid = cut",               "sólido = corte"),
        ["legend.dashedBend"]    = ("dashed = bend up",          "discontinuo = doblez"),
        ["legend.boldSpec"]      = ("highlighted = as specified","resaltado = según especificación"),
        ["legend.insideRadius"]  = ("inside radius Ri",          "radio interior Ri"),
        ["legend.fracInches"]    = ("dimensions in fractional inches","dimensiones en pulgadas fraccionarias"),
    };

    /// <summary>"English / Spanish" for a label id (falls back to the id if unknown).</summary>
    public static string T(string id) => Map.TryGetValue(id, out var v) ? $"{v.En} / {v.Es}" : id;

    public static string En(string id) => Map.TryGetValue(id, out var v) ? v.En : id;
    public static string Es(string id) => Map.TryGetValue(id, out var v) ? v.Es : id;

    /// <summary>
    /// The specified dimension's basis as a Spanish single word — Adentro (inside) / Afuera (outside),
    /// the shop's preferred terms. Deliberately Spanish-only (no ID/OD abbreviation, no English): it's
    /// the tag the shop reads straight off the dimension.
    /// </summary>
    public static string Basis(DimBasis b) => b == DimBasis.Inside ? "Adentro" : "Afuera";

    // ── Spanish material + part-type words for the bilingual title bar ───────
    private static readonly Dictionary<string, string> MaterialEsMap = new()
    {
        ["alum"]  = "Aluminio",          ["CRS"]  = "Acero laminado en frío",
        ["HRS"]   = "Acero laminado en caliente", ["galv"] = "Acero galvanizado",
        ["HRPO"]  = "Acero HRPO",         ["SS"]   = "Acero inoxidable",
        ["Brass"] = "Latón",             ["Copper"] = "Cobre",
    };

    /// <summary>Spanish display name for a material token (e.g. "CRS" → "Acero laminado en frío").</summary>
    public static string MaterialEs(string token) => MaterialEsMap.TryGetValue(token, out var v) ? v : token;

    private static readonly Dictionary<string, string> MaterialEnMap = new()
    {
        ["alum"]  = "Aluminum",          ["CRS"]  = "Cold Rolled Steel",
        ["HRS"]   = "Hot Rolled Steel",  ["galv"] = "Galvanized Steel",
        ["HRPO"]  = "HRPO Steel",        ["SS"]   = "Stainless Steel",
        ["Brass"] = "Brass",             ["Copper"] = "Copper",
    };

    /// <summary>English display name for a material token (e.g. "CRS" → "Cold Rolled Steel").</summary>
    public static string MaterialEn(string token) => MaterialEnMap.TryGetValue(token, out var v) ? v : token;

    /// <summary>Spanish part-type word for the title-bar subtitle.</summary>
    public static string TypeEs(PartType t) => t switch
    {
        PartType.UChannel    => "Canal U",
        PartType.LAngle      => "Ángulo",
        PartType.ZChannel    => "Canal Z",
        PartType.Pan         => "Bandeja",
        PartType.FlitchPlate => "Placa flitch",
        PartType.BasePlate   => "Placa base",
        PartType.PaddleBlind => "Disco ciego",
        PartType.Column      => "Columna",
        _ => t.ToString(),
    };
}
