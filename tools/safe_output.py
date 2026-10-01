"""Transactional output: validate a sibling temporary file before publishing it."""
import contextlib
import os
import pathlib
import tempfile


@contextlib.contextmanager
def staged_output(src, dst):
    src, dst = pathlib.Path(src).resolve(), pathlib.Path(dst).resolve()
    if src == dst or (dst.exists() and os.path.samefile(src, dst)):
        raise ValueError("refusing to overwrite the original")
    dst.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=".document-", suffix=dst.suffix, dir=dst.parent)
    os.close(fd)
    temp = pathlib.Path(name)
    try:
        yield temp
        # Recheck aliases immediately before atomic publication.
        if dst.exists() and os.path.samefile(src, dst):
            raise ValueError("refusing to overwrite the original")
        os.replace(temp, dst)
    finally:
        temp.unlink(missing_ok=True)
