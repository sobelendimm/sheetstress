from dataclasses import asdict, dataclass


@dataclass
class Finding:
    check: str
    severity: str  # "error" | "warning" | "info" | "skipped"
    location: str
    message: str

    def as_dict(self):
        return asdict(self)
