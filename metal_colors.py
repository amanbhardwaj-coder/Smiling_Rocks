"""Client color names -> canonical variation codes, without inferring availability."""
import re

COLOR_NAMES={'W':'White','Y':'Yellow','R':'Rose','WY':'White Yellow','WR':'White Rose','YR':'Yellow Rose'}
_ALIASES={
    'W':'W','WG':'W','WHITE':'W',
    'Y':'Y','YG':'Y','YELLOW':'Y',
    'R':'R','RG':'R','ROSE':'R',
    'WY':'WY','YW':'WY','WHITEYELLOW':'WY','YELLOWWHITE':'WY',
    'WR':'WR','RW':'WR','WHITEROSE':'WR','ROSEWHITE':'WR',
    'YR':'YR','RY':'YR','YELLOWROSE':'YR','ROSEYELLOW':'YR',
}

def parse_color(value):
    text=str(value or '').strip().upper()
    text=re.sub(r'^(?:14|18)\s*K[\s-]*','',text)
    text=text.replace('GOLD','').replace('TWO TONE','').replace('TWOTONE','')
    key=re.sub('[^A-Z]','',text)
    if key not in _ALIASES:
        raise ValueError(f"Unknown metal color {value!r}. Use White, Yellow, Rose, White Yellow or White Rose; separate choices with commas.")
    return _ALIASES[key]

def parse_colors(value):
    tokens=[t.strip() for t in re.split(r'[,#;\n]+',str(value or '')) if t.strip()]
    return tuple(dict.fromkeys(parse_color(token) for token in tokens))
