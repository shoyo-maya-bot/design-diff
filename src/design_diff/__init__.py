"""design-diff — 設計書（Excel）の新旧差分と、Excel の API 定義書と OpenAPI の差分を、レビュー用のレポートにする."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("design-diff")
except PackageNotFoundError:
    __version__ = "0.0.0+local"
