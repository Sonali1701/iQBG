import get_chaps
import json
with open('out.json', 'w') as f:
    json.dump(get_chaps.out, f, indent=2)
