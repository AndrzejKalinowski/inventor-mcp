"""Sketch icindeki geometri nesneleri (cizgi, daire, ...) icin bellek-ici kayit.

Inventor COM nesneleri JSON'a serialize edilemedigi icin, olusturulan her
sketch entity'sine bir string id atanir ve bu modulde COM referansi tutulur.
Sunucu ayni process icinde calistigi surece bu id'ler tool cagrilari arasinda
gecerlidir. Sunucu yeniden baslarsa `inventor_sketch_list_entities` ile
sketch'teki mevcut entity'ler yeniden listelenip kayit tazelenebilir.
"""

_entities: dict[str, object] = {}
_counters: dict[str, int] = {}


def _prefix(doc_name: str, sketch_name: str) -> str:
    return f"{doc_name}:{sketch_name}"


def register(doc_name: str, sketch_name: str, kind: str, obj) -> str:
    """Bir COM entity'sini kaydeder ve ona ait id'yi dondurur."""
    prefix = _prefix(doc_name, sketch_name)
    idx = _counters.get(prefix, 0)
    entity_id = f"{prefix}:{kind}{idx}"
    _counters[prefix] = idx + 1
    _entities[entity_id] = obj
    return entity_id


def resolve(entity_id: str):
    """Bir entity id'sine karsilik gelen COM nesnesini dondurur."""
    try:
        return _entities[entity_id]
    except KeyError:
        raise LookupError(
            f"Bilinmeyen entity_id: '{entity_id}'. Once 'inventor_sketch_list_entities' ile listeleyin."
        )


def clear_sketch(doc_name: str, sketch_name: str) -> None:
    """Bir sketch'e ait tum kayitli entity'leri temizler (sketch yeniden olusturulurken)."""
    prefix = _prefix(doc_name, sketch_name) + ":"
    for key in [k for k in _entities if k.startswith(prefix)]:
        del _entities[key]
    _counters.pop(_prefix(doc_name, sketch_name), None)
