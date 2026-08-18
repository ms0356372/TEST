from abc import ABC, abstractmethod
from pathlib import Path
from typing import Callable, ClassVar
LogCallback = Callable[[str], None]; ProgressCallback = Callable[[int, int], None]
class BaseAnalysis(ABC):
    key: ClassVar[str]; name: ClassVar[str]; required_templates: ClassVar[tuple[str,...]]
    required_headers: ClassVar[tuple[str,...]]; output_headers: ClassVar[tuple[str,...]]
    @abstractmethod
    def run(self, templates: dict[str,str], output_dir: str, log: LogCallback=lambda _:None, progress: ProgressCallback=lambda _a,_b:None) -> Path: ...
