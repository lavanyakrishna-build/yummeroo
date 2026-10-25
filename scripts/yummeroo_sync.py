"""Yummeroo recipe sync.

Rebuilds the recipe data inside the live Yummeroo app page from Muse's spreadsheet
(Yummeroo Recipes.xlsx). Used by the scheduled sync task.

Usage:
  python3 -I yummeroo_sync.py --xlsx recipes.xlsx --page live.html --modified 2026-10-16T01:42:50Z --out new.html

Notes copy (Lavanya's notes from the app to the "Yummeroo Notes" sheet):
  python3 -I yummeroo_sync.py notes --doc notes.json --sheet-rows 3 --out values.json

Recipe sync prints one JSON line. "status" is one of:
  unchanged  - the spreadsheet hasn't changed since the page was last built; nothing written
  no_changes - the spreadsheet changed but the recipes come out identical; nothing written
  updated    - new page written to --out; publish it
  abort      - something looks wrong (e.g. far fewer recipes than before); nothing written
"""
import argparse
"""Ingredient clean-up rules.

Each ingredient line becomes {raw, name, optional} where `name` is a
canonical ingredient (so "curd", "dahi" and "yogurt" all become "yogurt").
"""
import json
import datetime
import re
import sys
import collections
import openpyxl


# canonical name -> aliases (matched as whole words, longest alias wins)
CANON = {
    # --- produce ---
    "garlic": ["garlic", "garlic cloves", "garlic clove", "garlic bulb", "head of garlic", "roasted garlic", "garlic confit"],
    "onion": ["onion", "onions", "red onion", "red onions", "white onion", "yellow onion", "yellow onions", "pickled onions", "pickled red onions", "fried onions"],
    "shallot": ["shallot", "shallots"],
    "spring onion": ["spring onion", "spring onions", "scallion", "scallions", "green onion", "green onions", "spring onion greens", "spring onion whites", "chives"],
    "ginger": ["ginger", "ginger paste", "fresh ginger", "julienned ginger"],
    "ginger garlic paste": ["ginger garlic paste", "ginger-garlic paste"],
    "tomato": ["tomato", "tomatoes", "cherry tomatoes", "grape tomatoes", "roma tomatoes"],
    "canned tomatoes": ["canned whole tomatoes", "tomato sauce", "tomato puree", "tomato purée", "chopped tomatoes"],
    "tomato paste": ["tomato paste"],
    "sun-dried tomatoes": ["sun-dried tomatoes", "sun dried tomatoes"],
    "potato": ["potato", "potatoes", "roasted potatoes", "boiled potato", "boiled potatoes"],
    "sweet potato": ["sweet potato", "sweet potatoes"],
    "spinach": ["spinach", "palak", "chopped palak"],
    "cabbage": ["cabbage", "green cabbage", "hispi cabbage", "cone cabbage", "whole cabbage"],
    "cauliflower": ["cauliflower", "cauliflower florets"],
    "broccoli": ["broccoli", "broccoli florets", "broccolini"],
    "carrot": ["carrot", "carrots", "grated carrot", "grated carrots"],
    "bell pepper": ["bell pepper", "bell peppers", "capsicum", "red pepper", "red bell pepper", "yellow bell pepper", "green bell pepper", "roasted red pepper", "tricolour bell peppers"],
    "green chilli": ["green chilli", "green chillies", "green chili", "green chilies", "chillies", "chilies", "chilli", "chili", "fresh chillies", "red chilli pepper", "red chilli", "bird's eye red chilli", "small chillies", "serrano pepper", "chilli pepper"],
    "jalapeño": ["pickled jalapenos", "jalapenos", "jalapeños", "jalapeno", "jalapeño", "jalapeño rings", "pickled jalapeno", "pickled jalapeños", "jalapeno rings"],
    "dried red chilli": ["dried red chillies", "dried red chilli", "dried red chilies", "dry red chillies", "dried red chiles"],
    "mushroom": ["mushroom", "mushrooms", "button mushrooms", "oyster mushrooms", "cremini", "mushroom caps", "shiitake", "shiitake mushrooms", "dried mushrooms", "dried shiitake mushrooms", "porcini", "dried porcini mushrooms"],
    "zucchini": ["zucchini", "courgette"],
    "okra": ["okra", "fresh okra", "vendakkai", "bhindi", "lady finger", "ladies finger"],
    "eggplant": ["eggplant", "thai eggplant", "aubergine"],
    "cucumber": ["cucumber", "cucumbers"],
    "lettuce": ["lettuce", "romaine lettuce", "romaine", "lettuce leaves"],
    "arugula": ["arugula", "rocket", "rocket leaves"],
    "avocado": ["avocado"],
    "corn": ["corn", "sweet corn", "corn kernels", "baby corn"],
    "peas": ["peas", "frozen peas"],
    "edamame": ["edamame"],
    "green beans": ["green beans", "beans"],  # 'beans' alone alias resolved after specific beans
    "asparagus": ["asparagus"],
    "lotus stem": ["lotus stem"],
    "beetroot": ["beetroot", "beet"],
    "celery": ["celery", "celery sticks"],
    "lemon": ["lemon", "lemon juice", "lemon zest", "lemons"],
    "lime": ["lime", "lime juice", "limes", "lime wedges", "juice of 1 lime"],
    "lemongrass": ["lemongrass"],
    "galangal": ["galangal"],
    "kaffir lime leaves": ["kaffir lime leaves"],
    "curry leaves": ["curry leaves"],
    "mixed vegetables": ["mixed vegetables", "summer vegetables", "fajita veg"],
    "asian pear": ["asian pear", "fuji apple"],
    "coconut": ["coconut", "grated coconut", "coconut pieces"],
    # --- herbs ---
    "cilantro": ["cilantro", "coriander leaves", "fresh coriander", "chopped coriander", "coriander", "cilantro stems", "cilantro leaves"],
    "parsley": ["parsley", "parsley leaves"],
    "basil": ["basil", "basil leaves", "thai basil", "thai basil leaves", "fresh basil"],
    "mint": ["mint", "mint leaves", "fresh mint"],
    "dill": ["dill", "fresh dill"],
    "thyme": ["thyme", "fresh thyme", "dried thyme"],
    "rosemary": ["rosemary"],
    "oregano": ["oregano", "dried oregano"],
    "mixed herbs": ["italian seasoning", "mixed herbs", "italian herbs", "dried herbs", "herbs", "fresh herbs", "chopped herbs", "dried mixed italian seasoning", "herb"],
    "bay leaf": ["bay leaf", "bay leaves"],
    "kasuri methi": ["kasuri methi", "dried methi"],
    # --- dairy / eggs ---
    "egg": ["egg", "eggs", "egg whites", "liquid egg whites", "large eggs"],
    "butter": ["butter", "salted butter", "unsalted butter", "garlic butter", "garlic herb butter", "softened butter"],
    "ghee": ["ghee"],
    "milk": ["milk"],
    "cream": ["cream", "heavy cream", "heavy whipping cream", "fresh cream", "crème fraîche", "creme fraiche", "sour cream", "malai"],
    "yogurt": ["yogurt", "yoghurt", "greek yogurt", "greek yoghurt", "curd", "curds", "dahi", "hung curd", "greek-style yoghurt", "0% greek yogurt", "0% fat greek yogurt", "buttermilk", "plain yogurt", "thick yogurt"],
    "paneer": ["paneer", "low fat paneer", "low-fat paneer", "malai paneer", "low-fat high-protein paneer"],
    "cottage cheese": ["cottage cheese", "small-curd cottage cheese"],
    "cream cheese": ["cream cheese"],
    "mozzarella": ["mozzarella", "mozzarella cheese", "shredded mozzarella", "fat-free mozzarella"],
    "parmesan": ["parmesan", "parmesan cheese", "grated parmesan"],
    "cheddar": ["cheddar", "cheddar cheese", "sharp cheddar", "grated cheddar", "shredded cheddar cheese"],
    "feta": ["feta", "feta cheese"],
    "cheese": ["cheese", "cheese slices", "cheese slice", "processed cheese", "cheese cubes", "grated cheese", "shredded cheese", "mexican cheese blend", "shredded white cheese", "flavoured cheese", "shredded mexican cheese blend", "queso"],
    "mayonnaise": ["mayonnaise", "mayo", "light mayo"],
    # --- plant protein ---
    "tofu": ["tofu", "firm tofu", "extra firm tofu", "super firm tofu", "medium firm tofu", "smoked tofu", "silken tofu", "soft silken tofu", "silken-firm tofu", "tofu paste", "seasoned tofu crumble"],
    "tempeh": ["tempeh"],
    "seitan": ["seitan", "vital wheat gluten"],
    "soya chunks": ["soya chunks", "soy chunks"],
    "chickpeas": ["chickpeas", "chana"],
    "black beans": ["black beans"],
    "kidney beans": ["kidney beans", "rajma"],
    "butter beans": ["butter beans", "butterbeans", "cannellini"],
    "mixed beans": ["mixed beans", "refried beans"],
    "red lentils": ["red lentils", "masoor dal", "dried red lentils"],
    "lentils (dal)": ["toor dal", "tuvar dal", "arhar dal", "moong dal", "yellow moong dal", "chana dal", "urad dal", "split yellow lentils", "roasted chana dal"],
    "peanuts": ["peanuts", "crushed peanuts"],
    "cashews": ["cashews", "soaked cashews"],
    "walnuts": ["walnuts"],
    "sesame seeds": ["sesame seeds", "toasted sesame seeds", "white sesame seeds", "black and white sesame seeds", "roasted sesame seeds", "ground sesame seeds"],
    "watermelon seeds": ["watermelon seeds", "hulled watermelon seeds"],
    "peanut butter": ["peanut butter", "smooth peanut butter", "crunchy peanut butter", "unsweetened peanut butter"],
    "tahini": ["tahini", "sesame paste", "runny tahini"],
    "nutritional yeast": ["nutritional yeast"],
    # --- meat / fish ---
    "chicken": ["chicken", "chicken breast", "chicken mince", "ground chicken"],
    "ground beef": ["ground beef", "beef", "sauteed ground beef"],
    "ground pork": ["ground pork", "pork"],
    "ground meat": ["ground meat", "minced meat", "cooked ground meat"],
    "pepperoni": ["pepperoni"],
    # --- grains / carbs ---
    "rice": ["rice", "cooked rice", "basmati rice", "white rice", "steamed rice", "long grain rice", "uncooked rice"],
    "pasta": ["pasta", "spaghetti", "fusilli", "paccheri", "paccheri pasta", "pasta nests", "wide flat pasta nests", "large tube pasta", "lasagne sheets", "lasagna sheets", "wide pasta", "cooked pasta", "pasta of choice"],
    "noodles": ["noodles", "ramen noodles", "dry ramen noodles", "instant ramen noodles", "udon noodles", "chinese noodles", "shirataki noodles", "cooked noodles", "noodles of choice", "boiled noodles"],
    "bread": ["bread", "sourdough", "sourdough bread", "bread slices", "white bread slices", "toast", "ciabatta", "ciabattas", "french bread", "baguette", "toasted baguette", "garlic baguette", "multigrain or sourdough bread", "sourdough slices", "bread dough", "bread/roll dough", "bread dough/loaf"],
    "buns / rolls": ["burger buns", "burger bun", "buns", "buttered buns", "rolls", "hawaiian sweet rolls", "pav", "soft rolls", "bread or rolls"],
    "tortillas": ["tortillas", "tortilla", "tortilla wraps", "flour tortilla", "flour tortillas", "corn or flour tortillas", "lavash wraps", "low-carb taco-size tortillas", "low-calorie 10-inch tortillas", "small tortillas", "6-inch tortillas"],
    "roti / paratha": ["roti", "rotis", "paratha", "cooked rotis", "leftover roti"],
    "pita / flatbread": ["pita", "flatbread", "pita / flatbread"],
    "rice paper": ["rice paper", "rice paper sheets", "rice paper sheet"],
    "dumpling wrappers": ["dumpling wrappers", "wonton wrappers", "wonton or dumpling wrappers", "dumpling/wonton wrappers"],
    "tortilla chips": ["tortilla chips", "high-protein tortilla chips"],
    "flour": ["flour", "all-purpose flour", "all purpose flour", "plain flour", "whole wheat flour", "khapli wheat flour"],
    "chickpea flour": ["chickpea flour", "gram flour", "besan", "toasted besan"],
    "jowar flour": ["jowar flour"],
    "cornstarch": ["cornstarch", "corn starch", "cornflour", "corn flour", "potato starch", "corn flour slurry", "cornflour slurry"],
    "breadcrumbs": ["breadcrumbs", "panko", "panko breadcrumbs"],
    "cornflakes": ["cornflakes", "crushed cornflakes"],
    "potato sticks": ["potato sticks"],
    "flaxmeal": ["flaxmeal"],
    "baking powder": ["baking powder"],
    # --- sauces / condiments ---
    "soy sauce": ["dark sauce", "soy sauce/seasoning", "soy sauce", "light soy sauce", "dark soy sauce", "low sodium soy sauce", "tamari"],
    "oyster sauce": ["oyster sauce"],
    "vegetarian oyster sauce": ["vegetarian oyster sauce"],
    "fish sauce": ["fish sauce"],
    "vegan fish sauce": ["vegan fish sauce"],
    "hoisin sauce": ["hoisin sauce"],
    "gochujang": ["gochujang"],
    "gochugaru": ["gochugaru"],
    "chilli crisp / chilli oil": ["chili crisp", "chilli crisp", "chili oil", "chilli oil", "red chili oil", "chili crisp oil", "hot oil", "hot sesame oil"],
    "hot sauce": ["hot sauce", "tabasco", "sriracha", "buffalo sauce", "creamy hot sauce", "chilli sauce", "chili sauce"],
    "ketchup": ["ketchup", "tomato ketchup"],
    "mustard": ["mustard", "dijon mustard", "wholegrain mustard", "whole-grain mustard", "whole grain mustard", "mustard sauce"],
    "vinegar": ["vinegar", "rice vinegar", "rice wine vinegar", "white wine vinegar", "white vinegar", "black vinegar", "apple cider vinegar", "sherry", "red wine vinegar"],
    "mirin": ["mirin"],
    "worcestershire sauce": ["worcestershire sauce", "worcestershire", "vegan worcestershire", "vegan japanese worcestershire sauce"],
    "miso": ["miso", "miso paste"],
    "harissa": ["harissa", "harissa paste"],
    "chipotle": ["chipotle", "chipotle paste", "chipotle pepper in adobo", "chipotles", "diced chipotles", "chipotle flakes", "chipotle powder"],
    "curry paste": ["red curry paste", "curry paste"],
    "enchilada sauce": ["enchilada sauce", "red enchilada sauce"],
    "pesto": ["pesto", "basil pesto"],
    "green chutney": ["green chutney"],
    "salsa": ["salsa"],
    "pickles": ["pickles", "diced pickles"],
    "olives": ["olives", "green olives"],
    "tamarind": ["tamarind"],
    "ketjap manis": ["ketjap manis", "sweet soy sauce"],
    "coconut milk": ["coconut milk", "coconut cream", "canned coconut milk"],
    "stock / broth": ["stock", "broth", "vegetable stock", "veggie stock", "veg stock", "chicken stock", "chicken broth", "vegetable broth", "veggie broth", "mushroom stock", "stock cube", "veggie stock cube", "vegetable stock cube", "bouillon", "chicken bouillon powder", "vegan beef bouillon", "light stock", "kombu dashi powder"],
    "honey": ["honey"],
    "jaggery": ["jaggery", "gur", "vellam"],
    "maple syrup": ["maple syrup", "agave", "agave syrup"],
    # --- spices ---
    "cumin": ["cumin", "cumin seeds", "ground cumin", "cumin powder", "jeera", "jeera powder", "roasted cumin powder", "whole cumin"],
    "coriander powder": ["coriander powder", "ground coriander", "coriander seeds"],
    "turmeric": ["turmeric", "turmeric powder", "haldi"],
    "red chilli powder": ["red chilli powder", "chilli powder", "chili powder", "red chili powder", "kashmiri red chilli powder", "kashmiri red chili powder", "kashmiri chilli", "cayenne", "cayenne pepper"],
    "chilli flakes": ["chilli flakes", "chili flakes", "red chili flakes", "red chilli flakes", "red pepper flakes", "pepper flakes", "crushed red chilli flakes"],
    "paprika": ["paprika", "smoked paprika", "paprika powder", "smoked paprika powder"],
    "garam masala": ["garam masala", "biryani masala", "tandoori masala", "tandoori masala powder", "annas masala"],
    "chaat masala": ["chaat masala"],
    "curry powder": ["curry powder"],
    "garlic powder": ["garlic powder", "garlic granules"],
    "onion powder": ["onion powder"],
    "five spice": ["five spice powder", "chinese 5 spice", "chinese 5 spice powder"],
    "mustard seeds": ["mustard seeds"],
    "fenugreek seeds": ["fenugreek seeds", "methi seeds", "vendhayam"],
    "sambar powder": ["sambar masala", "sambar powder", "rasam powder"],
    "fennel seeds": ["fennel seeds"],
    "whole spices": ["cardamom", "green cardamom", "green cardamoms", "black cardamom", "cloves", "cinnamon", "dalchini", "mace", "black peppercorns", "peppercorns"],
    "asafoetida (hing)": ["hing", "asafoetida"],
    "saffron": ["saffron"],
    "msg": ["msg", "aromat powder"],
    "spices & seasonings": ["chicken spice", "spices", "spice mix", "spice blend", "mediterranean spice mix", "seasoning blend", "seasonings", "seasoning", "dry seasonings and spices", "ramen seasoning packet", "ferrari masala", "spice seasoning"],
    "taco seasoning": ["taco seasoning", "fajita seasoning", "chilli lime seasoning"],
    "mustard powder": ["mustard powder"],
    "cocoa powder": ["cocoa powder"],
    "liquid smoke": ["liquid smoke"],
    "black salt": ["black salt"],
    "white pepper": ["white pepper", "white pepper powder"],
    # --- staples ---
    "salt": ["salt", "sea salt", "smoked sea salt", "maldon smoked sea salt", "garlic salt", "fine sea salt"],
    "black pepper": ["black pepper", "black peppers", "pepper", "freshly ground black pepper", "black pepper powder", "crushed black pepper", "fresh cracked black pepper"],
    "oil": ["oil", "neutral oil", "vegetable oil", "cooking oil", "avocado oil", "groundnut oil", "cooking oil spray", "oil spray", "sesame oil", "toasted sesame oil", "mustard oil", "cold-pressed yellow mustard oil", "coconut oil"],
    "olive oil": ["olive oil", "extra virgin olive oil"],
    "sugar": ["sugar", "brown sugar", "coconut sugar", "granulated sugar"],
    "water": ["water", "pasta water", "reserved pasta water", "starchy pasta cooking water", "noodle water", "hot water", "cold water", "ice cubes", "mushroom soaking water"],
    # --- vegan stand-ins (kept separate so diet tags stay right) ---
    "vegan mayo": ["vegan mayo"],
    "vegan parmesan": ["vegan parmesan", "grated vegan parmesan"],
    "vegan feta": ["vegan feta"],
    "plant-based cream": ["plant cream", "plant-based cream", "plant-based yoghurt"],
    "vegan butter": ["vegan garlic butter"],
}

STAPLES = {"salt", "black pepper", "oil", "olive oil", "water", "sugar"}

MEAT = {"chicken", "ground beef", "ground pork", "ground meat", "pepperoni", "fish sauce", "oyster sauce"}
DAIRY = {"butter", "ghee", "milk", "cream", "yogurt", "paneer", "cottage cheese", "cream cheese",
         "mozzarella", "parmesan", "cheddar", "feta", "cheese"}
EGG = {"egg", "mayonnaise"}

OPTIONAL_PAT = re.compile(
    r"\boptional\b|\bfor garnish\b|\bto garnish\b|\bgarnish\b|\bto serve\b|\bfor serving\b|\bserving\b|"
    r"\bto taste\b|\bfor dipping\b|\bfor the pan\b|"
    r"\bfor frying\b|\bfor deep-frying\b|\bfor brushing\b|\bfor drizzling\b|\bfor searing\b|\bfor pan-frying\b")

ALIAS = []
for canon, aliases in CANON.items():
    for a in set(aliases + [canon]):
        ALIAS.append((a.lower(), canon))
ALIAS.sort(key=lambda x: -len(x[0]))

# lowercase-word patterns
LOW = {"cloves", "pepper", "chilli", "chili", "chillies", "chilies", "beans", "cheese", "herbs", "herb",
       "oil", "water", "stock", "broth", "toast", "rolls", "cream", "coriander", "salt", "sugar", "bread",
       "flour", "rice", "pasta", "noodles", "vinegar", "mustard", "tortilla", "cumin", "paprika"}
ALIAS.sort(key=lambda x: (x[0] in LOW, -len(x[0])))
ALIAS_RE = [(re.compile(r"(?<![a-z])" + re.escape(a) + r"(?![a-z'])"), c) for a, c in ALIAS]

unmatched = collections.Counter()


def clean_line(line):
    s = line.strip()
    # drop a short section label like "Burger sauce:" / "Filling:"
    m = re.match(r"^([A-Za-z][A-Za-z '\-/]{1,30}):\s+(.*)$", s)
    if m and not re.search(r"\d", m.group(1)):
        s = m.group(2)
    return s


def canon_of(text):
    low = text.lower()
    low = low.replace("’", "'")
    # remove parentheticals for matching, but keep a copy for optional check
    core = re.sub(r"\([^)]*\)", " ", low)
    # cut prep text after first comma
    first = core.split(",")[0]
    for chunk in (first, core):
        for rx, c in ALIAS_RE:
            if rx.search(chunk):
                return c
    return None


def split_top(s):
    """Split on ';' that are not inside parentheses."""
    out, depth, cur = [], 0, ""
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        if ch == ";" and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


VAGUE = re.compile(r"spice|seasoning|sauce|dip|batter|dressing|base|protein|liquid")


def parse_ingredients(cell):
    out = []
    if not cell:
        return out
    seen = set()
    for raw in split_top(str(cell)):
        raw = raw.strip()
        if not raw:
            continue
        line = clean_line(raw)
        # "Juice of 1 lime" etc handled by alias search
        # split compound entries like "paprika, fajita seasoning, cumin, oregano"
        parts = [line]
        core = re.sub(r"\([^)]*\)", "", line)
        if " - " in core and not canon_of(core.split(" - ", 1)[0]):  # "homemade guac - avocado, lime juice, ..."
            parts = [core.split(" - ", 1)[1]]
        listy = parts[0] is not line or (core.count(",") >= 3 and not re.search(r"\d", core))
        pieces = [p.strip() for p in re.split(r",|&|\band\b|/", parts[0])] if listy else parts
        optional = bool(OPTIONAL_PAT.search(raw.lower()))
        names = []
        for p in pieces:
            c = canon_of(p)
            if c and c not in names:
                names.append(c)
        if not names:
            c = canon_of(line)
            if c:
                names = [c]
        if not names:
            # fall back to a trimmed name
            fb = re.sub(r"\([^)]*\)", "", line.lower()).split(",")[0]
            fb = re.sub(r"^[\d\s/.\-½¼¾~]+", "", fb)
            fb = re.sub(r"^(cups?|tbsp|tsp|g|kg|ml|oz|lb|inch(es)?|cloves?|handfuls?|pinch( of)?|bunch|block|head|can|tin|pack|cube|a|an|of|small|medium|large|big)\b\s*", "", fb).strip()
            unmatched[fb] += 1
            names = [fb] if fb else []
            if VAGUE.search(fb):
                optional = True
        for n in names:
            if n == "spices & seasonings":
                optional = True
            out.append({"raw": raw, "name": n, "optional": optional})
    # collapse duplicates: an ingredient is required if any line needs it
    merged = collections.OrderedDict()
    for i in out:
        if i["name"] in merged:
            merged[i["name"]]["optional"] = merged[i["name"]]["optional"] and i["optional"]
            if i["raw"] not in merged[i["name"]]["raws"]:
                merged[i["name"]]["raws"].append(i["raw"])
        else:
            merged[i["name"]] = {"name": i["name"], "optional": i["optional"], "raws": [i["raw"]]}
    return list(merged.values())


CUISINE_MAP = {
    "asian fusion": (["Asian"], True), "chinese fusion": (["Chinese"], True),
    "mexican-indian fusion": (["Mexican", "Indian"], True), "indian fusion": (["Indian"], True),
    "indian-middle eastern fusion": (["Indian", "Middle Eastern"], True), "chinese-inspired": (["Chinese"], False),
    "chinese-american fusion": (["Chinese", "American"], True), "indo-chinese": (["Indian", "Chinese"], True),
    "italian-american": (["Italian", "American"], True), "italian fusion": (["Italian"], True),
    "korean fusion": (["Korean"], True), "southeast asian": (["Southeast Asian"], False),
    "fusion": (["Fusion"], True), "vegan": (["Other"], False), "vegetarian": (["Other"], False),
}


def split_steps(text):
    """Turn Muse's method text into a list of steps."""
    text = (text or "").strip()
    if not text:
        return []
    if "\n" in text:
        parts = [re.sub(r"^\s*\d+[.)]\s*", "", p).strip() for p in re.split(r"\n+", text)]
    else:
        parts = [p.strip() for p in re.split(r"(?<=[.!?])\s+(?=[A-Z(])", text)]
    steps = []
    for p in parts:
        if not p:
            continue
        if steps and len(p) < 25:
            steps[-1] += " " + p
        else:
            steps.append(p)
    return steps


def clean_text(s):
    if s is None:
        return ""
    s = str(s).replace("\\_", "_").replace("\\#", "#").replace("\\[", "[").replace("\\]", "]").replace("\\~", "~")
    return s.strip()


def build_data(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    header = [c.value for c in ws[1]]
    recipes = []
    for idx, cells in enumerate(ws.iter_rows(min_row=2)):
        row = [c.value for c in cells]
        # A link cell may show text like "View reel on Instagram" with the real URL behind it.
        for i, c in enumerate(cells):
            if header[i] == "Post link" and c.hyperlink and c.hyperlink.target and not str(c.value or "").startswith("http"):
                row[i] = c.hyperlink.target
        r = dict(zip(header, row))
        if not r.get("Name"):
            continue
        name = clean_text(r["Name"])
        cz = clean_text(r.get("Cuisine"))
        cuisines, fusion = CUISINE_MAP.get(cz.lower(), ([cz or "Other"], False))
        notes = clean_text(r.get("Notes"))
        incomplete = None
        m = re.search(r"\[INCOMPLETE DATA\s*[—-]\s*(.*?)\]", notes)
        if m:
            incomplete = m.group(1).strip()
            notes = (notes[:m.start()] + notes[m.end():]).strip()
        ingredients = parse_ingredients(clean_text(r.get("Key ingredients")))
        raw_all = clean_text(r.get("Key ingredients")).lower()
        no_qty = raw_all.count("qty not shown") + raw_all.count("unmeasured")
        if not incomplete and ingredients and no_qty >= max(3, len(ingredients) * 0.6):
            incomplete = "Most quantities aren't given in the post."
        names = {i["name"] for i in ingredients if not i["optional"]}
        rawreq = " ".join(" ".join(i["raws"]) for i in ingredients if not i["optional"]).lower()
        if re.search(r"chicken (stock|broth|bouillon)", rawreq):
            names.add("chicken")
        has_meat = bool(names & MEAT)
        has_egg = "egg" in names or "mayonnaise" in names
        has_dairy = bool(names & DAIRY)
        if "honey" in names:
            pass
        diet = "meat" if has_meat else ("vegan" if not (has_egg or has_dairy) else "vegetarian")
        author = clean_text(r.get("Author"))
        if author and not author.startswith("@") and author != "AI-generated":
            author = "@" + author
        t = r.get("Time")
        recipes.append({
            "id": idx,
            "name": name,
            "cuisines": cuisines,
            "fusion": fusion,
            "meal": clean_text(r.get("Meal type")) or "Other",
            "time": int(t) if isinstance(t, (int, float)) else None,
            "difficulty": clean_text(r.get("Difficulty")) or None,
            "diet": diet,
            "egg": has_egg,
            "author": author,
            "link": clean_text(r.get("Post link")) or None,
            "notes": notes,
            "incomplete": incomplete,
            "tried": bool(r.get("Tried")),
            "ai": bool(r.get("AI-made")),
            "ingredients": ingredients,
            "method": split_steps(clean_text(r.get("Method"))),
        })
    aliases = {a: c for a, c in ALIAS if a != c and len(a) > 2}
    return {"recipes": recipes, "staples": sorted(STAPLES), "aliases": aliases}



DATA_RE = re.compile(r'(<script type="application/json" id="data">)(.*?)(</script>)', re.S)


def sync(xlsx, page, modified, out):
    html = open(page, encoding="utf-8").read()
    m = DATA_RE.search(html)
    if not m:
        return {"status": "abort", "reason": "Couldn't find the recipe data block in the live page."}
    old = json.loads(m.group(2).replace("<\\/", "</"))
    if old.get("sourceModified") == modified:
        return {"status": "unchanged", "recipes": len(old.get("recipes", []))}
    new = build_data(xlsx)
    old_r, new_r = old.get("recipes", []), new["recipes"]
    if len(new_r) < max(10, int(len(old_r) * 0.5)):
        return {"status": "abort", "reason": f"The spreadsheet has {len(new_r)} recipes, down from {len(old_r)}. Not updating in case the file is broken."}
    key = lambda r: r.get("link") or r.get("name")
    strip = lambda r: {k: v for k, v in r.items() if k != "id"}
    old_by, new_by = {key(r): r for r in old_r}, {key(r): r for r in new_r}
    added = [new_by[k]["name"] for k in new_by if k not in old_by]
    removed = [old_by[k]["name"] for k in old_by if k not in new_by]
    changed = [new_by[k]["name"] for k in new_by if k in old_by and strip(new_by[k]) != strip(old_by[k])]
    new["sourceModified"] = modified
    new["syncedAt"] = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    summary = {"recipes": len(new_r), "added": added, "removed": removed, "changed": changed,
               "unmatched_ingredients": [k for k, _ in unmatched.most_common(20)]}
    if not (added or removed or changed or new["aliases"] != old.get("aliases") or new["staples"] != old.get("staples")):
        # record the new file time anyway, so the next run can stop early
        summary["status"] = "no_changes"
    else:
        summary["status"] = "updated"
    payload = json.dumps(new, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    open(out, "w", encoding="utf-8").write(html[:m.start(2)] + payload + html[m.end(2):])
    return summary


def notes_values(doc_path, sheet_rows):
    """Turn Lavanya's notes document from the app into the full contents of the Notes tab.

    The sheet is a mirror of the app: one row per note, sorted by recipe name; rows left
    over from earlier runs are blanked. Text that Sheets would read as a formula, number
    or date gets a leading apostrophe so it stays text.
    """
    d = json.load(open(doc_path, encoding="utf-8"))
    notes = d.get("notes") if isinstance(d.get("notes"), dict) else (d.get("data") or {}).get("notes") or {}
    def as_text(s):
        s = str(s or "")
        return "'" + s if s and (s[0] in "=+-@" or re.fullmatch(r"[\d.,/\s:-]+", s)) else s
    rows = []
    for key, n in notes.items():
        if not isinstance(n, dict) or not n.get("note"):
            continue
        rows.append([as_text(n.get("name") or key), n.get("link") or "", as_text(n.get("note")), as_text(n.get("updated") or "")])
    rows.sort(key=lambda r: r[0].lstrip("'").lower())
    values = [["Recipe", "Instagram link", "My notes", "Last updated"]] + rows
    total = max(len(values), sheet_rows)
    values += [["", "", "", ""]] * (total - len(values))
    return {"range": f"Notes!A1:D{total}", "values": values, "notes": len(rows)}


if __name__ == "__main__" and len(sys.argv) > 1 and sys.argv[1] == "notes":
    # python3 -I yummeroo_sync.py notes --doc notes.json --sheet-rows 5 --out values.json
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd")
    ap.add_argument("--doc", required=True)
    ap.add_argument("--sheet-rows", type=int, required=True, help="rows the Notes tab currently has filled, header included")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    v = notes_values(a.doc, a.sheet_rows)
    json.dump({"range": v["range"], "values": v["values"]}, open(a.out, "w", encoding="utf-8"), ensure_ascii=False)
    print(json.dumps({"status": "ok", "notes": v["notes"], "range": v["range"]}))
elif __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--page", required=True)
    ap.add_argument("--modified", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    print(json.dumps(sync(a.xlsx, a.page, a.modified, a.out), ensure_ascii=False))
