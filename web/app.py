"""
Flask web application for NoSQL Lab 11 – Sharding & Shard Key Analysis.
Lab 11.1: Shard Key Decision Matrix (design/analysis lab)
Lab 11.2: Shard Key Hotspot Detection (data distribution simulation)
"""

import os, sys, hashlib, random, time
from datetime import datetime, timedelta

_here = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_here, ".."))

from flask import Flask, render_template, request, jsonify
from pymongo import MongoClient, ASCENDING, HASHED
from pymongo.errors import PyMongoError
from bson import ObjectId

MONGO_URI = os.getenv("MONGO_URI", "mongodb://mongo:27017/")
NUM_DOCS   = 10_000

app = Flask(__name__)

# ── lazy singleton ────────────────────────────────────────────────────────────
_client = None

def _get_client():
    global _client
    if _client is None:
        _client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=15000)
    return _client

def _db(name="lab11"):
    return _get_client()[name]


# ═══════════════════════════════════════════════════════════════════════════════
# LAB 11.1 helpers – Shard Key Decision Matrix
# ═══════════════════════════════════════════════════════════════════════════════

SHARD_KEY_CANDIDATES = [
    {
        "name": "order_id",
        "cardinality": "Very High (millions of unique orders)",
        "cardinality_score": 4,
        "distribution": "Even (sequential IDs spread across shards with hashing)",
        "distribution_score": 3,
        "query_targeting": "Poor – queries by customer_id scatter-gather all shards",
        "query_targeting_score": 1,
        "monotonic": True,
        "hotspot_risk": "High – auto-increment IDs send all new writes to one chunk",
        "verdict": "AVOID (raw). Use hashed(order_id) if order-centric writes dominate.",
        "score": 2,
        "color": "warning",
    },
    {
        "name": "customer_id",
        "cardinality": "High (millions of customers)",
        "cardinality_score": 4,
        "distribution": "Potentially Uneven – power-law customers dominate chunks",
        "distribution_score": 2,
        "query_targeting": "Excellent – all orders for a customer hit one shard",
        "query_targeting_score": 4,
        "monotonic": False,
        "hotspot_risk": "Medium – VIP customers create large chunks (jumbo risk)",
        "verdict": "GOOD for customer-scoped reads. Watch for jumbo chunks on VIPs.",
        "score": 3,
        "color": "info",
    },
    {
        "name": "order_date",
        "cardinality": "Low-Medium (one value per second max)",
        "cardinality_score": 2,
        "distribution": "Very Uneven – newest chunk absorbs all current writes",
        "distribution_score": 1,
        "query_targeting": "Good for time-range queries (single shard for recent window)",
        "query_targeting_score": 3,
        "monotonic": True,
        "hotspot_risk": "Very High – monotonically increasing = permanent write hotspot",
        "verdict": "AVOID – classic hotspot pattern. Never use timestamps as shard key.",
        "score": 1,
        "color": "danger",
    },
    {
        "name": "hashed(customer_id)",
        "cardinality": "Very High (hash space is 2^64)",
        "cardinality_score": 5,
        "distribution": "Excellent – uniform distribution across all shards",
        "distribution_score": 5,
        "query_targeting": "Good – customer queries still target one shard via hash lookup",
        "query_targeting_score": 4,
        "monotonic": False,
        "hotspot_risk": "Very Low – hash randomises chunk ownership",
        "verdict": "RECOMMENDED – best balance of distribution + customer targeting.",
        "score": 5,
        "color": "success",
    },
]


# ═══════════════════════════════════════════════════════════════════════════════
# LAB 11.2 helpers – Hotspot Detection
# ═══════════════════════════════════════════════════════════════════════════════

def _simulate_hotspot(n=NUM_DOCS, num_shards=2):
    """
    Simulate monotonically-increasing timestamp shard key distribution.
    In a real sharded cluster all new docs go to the latest (highest) chunk,
    which always lives on one shard until a split/migrate occurs.
    We model this by putting the last 95% of docs on shard-0.
    """
    base = datetime(2025, 1, 1)
    docs = []
    for i in range(n):
        ts = base + timedelta(seconds=i)
        docs.append({"ts": ts.isoformat(), "user": f"u-{i % 500:04d}", "value": i})

    # Simulate chunk ownership: shard-0 owns [min, split_point), shard-1 owns [split_point, max)
    # Initial split is at the midpoint of the original range; all NEW inserts go to shard-1
    split_idx = int(n * 0.05)        # only early 5% ended up on shard-0 after migrates
    shard_counts = [split_idx, n - split_idx]

    buckets = []
    for s in range(num_shards):
        pct = shard_counts[s] / n * 100
        buckets.append({"shard": f"shard-{s}", "docs": shard_counts[s], "pct": round(pct, 1)})
    return buckets, docs[:5]


def _simulate_hashed(n=NUM_DOCS, num_shards=2):
    """
    Simulate hashed shard key (hashed user_id / random ObjectId) distribution.
    Hash uniformly distributes across the shard space.
    """
    shard_counts = [0] * num_shards
    sample_docs = []
    for i in range(n):
        uid = str(ObjectId())
        h = int(hashlib.md5(uid.encode()).hexdigest(), 16)
        shard_counts[h % num_shards] += 1
        if i < 5:
            sample_docs.append({"user_id": uid, "value": i})

    buckets = []
    for s in range(num_shards):
        pct = shard_counts[s] / n * 100
        buckets.append({"shard": f"shard-{s}", "docs": shard_counts[s], "pct": round(pct, 1)})
    return buckets, sample_docs


def _simulate_queries():
    """Return explain()-like query plan comparison between hotspot and hashed collections."""
    return {
        "hotspot_range": {
            "queryPlanner": {
                "winningPlan": {"stage": "SHARDING_FILTER", "shards": 1},
                "note": "Range query on ts → targets SINGLE shard (chunk range match). Efficient.",
            },
            "executionStats": {"nReturned": 365, "totalKeysExamined": 365, "shardsContacted": 1},
        },
        "hashed_range": {
            "queryPlanner": {
                "winningPlan": {"stage": "SHARD_MERGE", "shards": 2},
                "note": "Range query on hashed user_id → BROADCAST to ALL shards (scatter-gather). Inefficient for ranges.",
            },
            "executionStats": {"nReturned": 2000, "totalKeysExamined": 10000, "shardsContacted": 2},
        },
    }


def _mongo_insert_hotspot(n=NUM_DOCS):
    """Actually insert into MongoDB if available (falls back to simulation)."""
    try:
        col = _db()["hotspot_data"]
        col.drop()
        col.create_index([("ts", ASCENDING)])
        base = datetime(2025, 1, 1)
        batch = []
        for i in range(n):
            ts = base + timedelta(seconds=i)
            batch.append({"ts": ts, "user": f"u-{i % 500:04d}", "value": i})
            if len(batch) == 500:
                col.insert_many(batch)
                batch.clear()
        if batch:
            col.insert_many(batch)
        count = col.count_documents({})
        return {"ok": True, "inserted": count}
    except PyMongoError as e:
        return {"ok": False, "error": str(e)}


def _mongo_insert_hashed(n=NUM_DOCS):
    """Actually insert into MongoDB if available."""
    try:
        col = _db()["even_data"]
        col.drop()
        col.create_index([("user_id", HASHED)])
        batch = []
        for i in range(n):
            batch.append({"user_id": str(ObjectId()), "value": i})
            if len(batch) == 500:
                col.insert_many(batch)
                batch.clear()
        if batch:
            col.insert_many(batch)
        count = col.count_documents({})
        return {"ok": True, "inserted": count}
    except PyMongoError as e:
        return {"ok": False, "error": str(e)}


# ═══════════════════════════════════════════════════════════════════════════════
# Routes
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/lab1")
def lab1():
    return render_template("lab1.html", candidates=SHARD_KEY_CANDIDATES)

@app.route("/lab2")
def lab2():
    return render_template("lab2.html")


# ── Lab 11.1 API ─────────────────────────────────────────────────────────────

@app.route("/api/shard-keys")
def api_shard_keys():
    return jsonify(SHARD_KEY_CANDIDATES)


# ── Lab 11.2 API ─────────────────────────────────────────────────────────────

@app.route("/api/simulate/hotspot")
def api_simulate_hotspot():
    n = int(request.args.get("n", NUM_DOCS))
    buckets, sample = _simulate_hotspot(n)
    return jsonify({"buckets": buckets, "sample_docs": sample, "total": n, "type": "monotonic_timestamp"})

@app.route("/api/simulate/hashed")
def api_simulate_hashed():
    n = int(request.args.get("n", NUM_DOCS))
    buckets, sample = _simulate_hashed(n)
    return jsonify({"buckets": buckets, "sample_docs": sample, "total": n, "type": "hashed_user_id"})

@app.route("/api/simulate/queries")
def api_simulate_queries():
    return jsonify(_simulate_queries())

@app.route("/api/insert/hotspot", methods=["POST"])
def api_insert_hotspot():
    n = int(request.json.get("n", NUM_DOCS))
    t0 = time.perf_counter()
    result = _mongo_insert_hotspot(n)
    result["elapsed_s"] = round(time.perf_counter() - t0, 2)
    return jsonify(result)

@app.route("/api/insert/hashed", methods=["POST"])
def api_insert_hashed():
    n = int(request.json.get("n", NUM_DOCS))
    t0 = time.perf_counter()
    result = _mongo_insert_hashed(n)
    result["elapsed_s"] = round(time.perf_counter() - t0, 2)
    return jsonify(result)

@app.route("/api/stats")
def api_stats():
    """Return per-collection stats from MongoDB."""
    try:
        client = _get_client()
        client.admin.command("ping")
        db = _db()
        stats = {}
        for col_name in ["hotspot_data", "even_data"]:
            if col_name in db.list_collection_names():
                s = db.command("collStats", col_name)
                stats[col_name] = {
                    "count":      s.get("count", 0),
                    "size_bytes": s.get("size", 0),
                    "avg_obj":    s.get("avgObjSize", 0),
                }
            else:
                stats[col_name] = None
        return jsonify({"ok": True, "stats": stats})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/health")
def health():
    try:
        _get_client().admin.command("ping")
        return jsonify({"status": "ok", "mongo": "connected"})
    except Exception as e:
        return jsonify({"status": "degraded", "mongo": str(e)}), 200


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
