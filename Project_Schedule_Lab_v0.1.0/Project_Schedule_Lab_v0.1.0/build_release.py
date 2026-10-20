"""Package sources, scripts, examples and evidence; never package local databases or environments."""
from pathlib import Path
import hashlib
import zipfile

ROOT=Path(__file__).resolve().parent


def main():
    dist=ROOT/'dist';dist.mkdir(exist_ok=True)
    target=dist/'ProjectScheduleLab-0.2.2.zip'
    sources=sorted(ROOT.glob('*.py'))+sorted(ROOT.glob('*.md'))+sorted(ROOT.glob('requirements*.txt'))
    sources += [ROOT/n for n in ('install.bat','start.bat','install.sh','start.sh','common_Rev1.xlsx')]
    sources += sorted((ROOT/'examples').glob('*'))+sorted((ROOT/'reports').glob('*'))
    with zipfile.ZipFile(target,'w',compression=zipfile.ZIP_DEFLATED) as z:
        for path in sources:
            if path.is_file(): z.write(path,'ProjectScheduleLab-0.2.2/'+path.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(target) as z:
        if z.testzip() is not None: raise RuntimeError('Archive checksum verification failed')
    digest=hashlib.sha256(target.read_bytes()).hexdigest()
    (dist/(target.name+'.sha256')).write_text(digest+'  '+target.name+'\n',encoding='ascii')
    print(target);print(digest)


if __name__=='__main__': main()
