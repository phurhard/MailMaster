import time
from collections import defaultdict
from fastapi import HTTPException

class IdempotencyManager:
    """
    Simple in-memory manager to prevent duplicate concurrent or rapid-fire requests.
    """
    def __init__(self, expiry_seconds: int = 10):
        self.locks = {} # (user_id, action, resource_id) -> expiry_time
        self.expiry_seconds = expiry_seconds

    def check_and_lock(self, user_id: str, action: str, resource_id: str):
        key = (user_id, action, resource_id)
        now = time.time()
        
        # Clean up old locks
        self.locks = {k: v for k, v in self.locks.items() if v > now}
        
        if key in self.locks:
            return False
        
        self.locks[key] = now + self.expiry_seconds
        return True

# Global instance
idempotency_manager = IdempotencyManager()
