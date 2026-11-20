"""Shared MongoDB connection helpers for Lab 11."""
import os
from pymongo import MongoClient

MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017/")


def get_client():
    return MongoClient(MONGO_URI, serverSelectionTimeoutMS=10000)


def get_db(db_name: str):
    return get_client()[db_name]


def reset_collection(db_name: str, col_name: str):
    db = get_db(db_name)
    db[col_name].drop()
    return db[col_name]


def banner(msg: str):
    line = "=" * (len(msg) + 4)
    print(f"\n{line}\n  {msg}\n{line}")
