class Phase2BadRequestError(Exception):
    def __init__(self, code: str, message: str):
        self.code = str(code or "BAD_REQUEST")
        self.message = str(message or "")
        super().__init__(f"{self.code}: {self.message}")

