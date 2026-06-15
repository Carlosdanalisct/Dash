import json

try:
    from cachetools import TTLCache
except ModuleNotFoundError:
    TTLCache = None


if TTLCache:
    api_cache = TTLCache(maxsize=100, ttl=300)
else:
    api_cache = {}


def cache_available():
    return TTLCache is not None


def cache_key(name, query):
    compact = {key: values for key, values in sorted(query.items())}
    return f"{name}:{json.dumps(compact, ensure_ascii=False, sort_keys=True)}"


def get_cached(name, query, builder):
    key = cache_key(name, query)
    if key not in api_cache:
        api_cache[key] = builder()
    return api_cache[key]


def clear_api_cache():
    api_cache.clear()

