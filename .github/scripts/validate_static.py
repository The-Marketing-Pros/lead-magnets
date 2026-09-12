"""Check local static references without network requests or reporting page content."""
import argparse
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlsplit


class References(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.references = []

    def handle_starttag(self, tag, attrs):
        for key, value in attrs:
            if key in ('href', 'src') and value:
                self.references.append(value)


def validate(root, private=False):
    root = Path(root).resolve()
    failures = []
    if not (root / 'index.html').is_file():
        failures.append('Missing root index.html')
    pages = [p for p in root.rglob('*.html') if '.git' not in p.relative_to(root).parts]
    for page in pages:
        parser = References()
        parser.feed(page.read_text(encoding='utf-8'))
        for value in parser.references:
            url = urlsplit(value)
            if url.scheme or url.netloc or not url.path:
                continue
            path = unquote(url.path)
            target = ((root / path.lstrip('/')) if path.startswith('/') else (page.parent / path)).resolve()
            if not target.is_relative_to(root):
                failures.append(f'{page.relative_to(root)}: reference leaves the published root')
            elif not target.is_file() and not (target / 'index.html').is_file():
                failures.append(f'{page.relative_to(root)}: missing local target {path}')
    if private:
        headers = root / '_headers'
        text = headers.read_text().lower() if headers.exists() else ''
        required = ['/*', 'x-robots-tag: noindex, nofollow, noarchive', 'cache-control: private, no-store']
        if any(line not in text for line in required):
            failures.append('Missing required private/no-store or noindex headers')
    return len(pages), failures


if __name__ == '__main__':
    args = argparse.ArgumentParser()
    args.add_argument('root')
    args.add_argument('--private', action='store_true')
    options = args.parse_args()
    count, failures = validate(options.root, options.private)
    for failure in failures:
        print(f'::error::{failure}')
    print(f'Validated {count} HTML files; {len(failures)} reference/header failures. Live access control was not tested.')
    raise SystemExit(bool(failures))
