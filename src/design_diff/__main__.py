"""python -m design_diff（design-diff コマンドと同じ. PATH にスクリプトが無い環境でも使える）."""

from .cli.main import entrypoint

if __name__ == "__main__":
    entrypoint()
