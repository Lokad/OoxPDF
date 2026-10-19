"""Compare page-clipped inline glyph/frame origins with pinned Office measurements.

Requires PyMuPDF in the validation environment; never renders or rebuilds the library.
Numerical axes and thresholds stay explicit. Pixel/ink checks remain separate.
"""
import argparse
import hashlib
import json
import unicodedata
from pathlib import Path
import pymupdf

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()

def measure(path, following_text=None):
    glyphs, frames, following = [], [], []
    with pymupdf.open(path) as pdf:
        for page_index, page in enumerate(pdf):
            for block in page.get_text('rawdict')['blocks']:
                for line in block.get('lines', []):
                    if following_text is not None and ''.join(char['c'] for span in line['spans'] for char in span['chars']).strip() == following_text:
                        following.append({'Page': page_index + 1, 'X': line['spans'][0]['origin'][0], 'Y': line['spans'][0]['origin'][1]})
                    for span in line['spans']:
                        if 'CambriaMath' not in span['font'].replace('-', '').replace(' ', ''):
                            continue
                        for char in span['chars']:
                            if char['c'].strip():
                                glyphs.append({'Text': unicodedata.normalize('NFKC', char['c']), 'Page': page_index + 1,
                                    'X': char['origin'][0], 'Y': char['origin'][1], 'Size': span['size']})
            for drawing in page.get_drawings():
                rect = drawing['rect']
                if drawing['type'] == 's' and drawing['color'] == (0., 0., 0.) and rect.width > 100 and rect.height > 20:
                    frames.append({'Page': page_index + 1, 'Rect': list(rect), 'StrokeWidth': drawing['width'],
                        'InsidePage': rect in page.rect})
    return glyphs, frames, following

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--golden', type=Path, required=True)
    parser.add_argument('--candidate', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        raise ValueError('Use a fresh geometry report path')
    gold = json.loads(args.golden.read_text(encoding='utf-8-sig'))
    outcomes_path = args.candidate / 'outcomes.json'
    outcomes = json.loads(outcomes_path.read_text(encoding='utf-8-sig'))
    by_key = {(row['Id'], row['Mode']): row for row in outcomes['Rows']}
    jobs = []
    for row in gold['Rows']:
        key = row['Id'], row['Mode']
        candidate = by_key.get(key)
        path = args.candidate / (row['Id'] + '-' + row['Mode']) / 'candidate.pdf'
        if candidate is None or candidate['InputSha256'] != row['InputSha256'] or sha(path) != candidate['PdfSha256']:
            raise ValueError('Pinned candidate differs: ' + str(key))
        jobs.append((row, path))
    rows = []
    for row, path in jobs:
        expected_following = row.get('WordFollowingAnchor')
        math, frames, following = measure(path, row.get('FollowingText', 'After public math table.') if expected_following else None)
        expected, expected_frames = row['WordMathAnchors'], row['WordFrames']
        axes = row.get('GeometryAxes', ['X', 'Y'])
        result = {'Id': row['Id'], 'Mode': row['Mode'], 'PdfSha256': sha(path), 'WordPdfSha256': row['WordPdfSha256'],
            'GeometryAxes': axes, 'NativeMath': math, 'NativeFrames': frames,
            'MissingPageClippedMathGlyphs': len(expected) - len(math), 'Failures': []}
        if len(math) != len(expected) or any(a['Text'] != e['Text'] or a['Page'] != e['Page'] for a, e in zip(math, expected)):
            result['Failures'].append('MathCountOrSourceOrder')
        elif expected:
            for axis in ['X', 'Y', 'Size']:
                result['MaxMath' + axis + 'Error'] = max(abs(a[axis] - e[axis]) for a, e in zip(math, expected))
            result['Failures'] += ['Math' + axis for axis in axes if result['MaxMath' + axis + 'Error'] >= 1.]
            if result['MaxMathSizeError'] >= .1:
                result['Failures'].append('MathPhysicalSize')
        if len(frames) != len(expected_frames) or any(a['Page'] != e['Page'] for a, e in zip(frames, expected_frames)):
            result['Failures'].append('FrameCountOrPage')
        elif expected_frames:
            for axis, indices in [('X', [0, 2]), ('Y', [1, 3])]:
                result['MaxFrame' + axis + 'Error'] = max(abs(a['Rect'][i] - e['Rect'][i])
                    for a, e in zip(frames, expected_frames) for i in indices)
            result['MaxStrokeError'] = max(abs(a['StrokeWidth'] - e['StrokeWidth']) for a, e in zip(frames, expected_frames))
            result['Failures'] += ['Frame' + axis for axis in axes if result['MaxFrame' + axis + 'Error'] >= 1.]
            if result['MaxStrokeError'] >= .01:
                result['Failures'].append('StrokeWidth')
        if any(not frame['InsidePage'] for frame in frames):
            result['Failures'].append('OffPageFrame')
        if expected_following:
            result['NativeFollowing'] = following
            if len(following) != 1 or following[0]['Page'] != expected_following['Page']:
                result['Failures'].append('FollowingSourceCountOrPage')
            else:
                for axis in ['X', 'Y']:
                    result['Following' + axis + 'Error'] = following[0][axis] - expected_following[axis]
                result['FollowingResidualAxes'] = [axis for axis in ['X', 'Y'] if abs(result['Following' + axis + 'Error']) >= 1.]
                result['Failures'] += ['Following' + axis for axis in row.get('FollowingGeometryAxes', [])
                    if axis in result['FollowingResidualAxes']]
        rows.append(result)
    report = {'Scope': 'Page-clipped glyph origins, authored black frame edges and physical sizes. This does not prove visible ink or content inside PDF clips.',
        'GoldenSha256': sha(args.golden), 'OutcomesSha256': sha(outcomes_path), 'PyMuPdfVersion': pymupdf.VersionBind,
        'OriginThresholdPoints': 1., 'FontSizeThresholdPoints': .1, 'StrokeThresholdPoints': .01,
        'FailedOutcomes': sum(bool(row['Failures']) for row in rows),
        'FollowingResidualOutcomes': sum(bool(row.get('FollowingResidualAxes')) for row in rows), 'Rows': rows}
    args.output.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
    for row in rows:
        print(row['Id'], row['Mode'], {k: v for k, v in row.items() if k.startswith(('Max', 'Following')) or k == 'Failures'})
    return int(report['FailedOutcomes'] != 0)

if __name__ == '__main__':
    raise SystemExit(main())
