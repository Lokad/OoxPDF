"""Compare page-clipped inline glyph/frame origins with pinned Office measurements.

Requires PyMuPDF in the validation environment; never renders or rebuilds the library.
Numerical axes and thresholds stay explicit. Pixel/ink checks remain separate.
"""
import argparse
import hashlib
import json
import math
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

def measure_prose(path, prefixes):
    """Keep every prefix match: duplicate or missing source lines must fail."""
    if len(set(prefixes)) != len(prefixes):
        raise ValueError('Prose anchor prefixes must be unique')
    matches = {prefix: [] for prefix in prefixes}
    with pymupdf.open(path) as pdf:
        for page_index, page in enumerate(pdf):
            for block in page.get_text('rawdict')['blocks']:
                for line in block.get('lines', []):
                    characters = [(char, span['size']) for span in line['spans'] for char in span['chars']]
                    value = ' '.join(unicodedata.normalize('NFKC', ''.join(char['c'] for char, _ in characters)).split())
                    first = next(((char, size) for char, size in characters if char['c'].strip()), None)
                    if first is None:
                        continue
                    for prefix in prefixes:
                        if value.startswith(prefix):
                            char, size = first
                            matches[prefix].append({'Text': prefix, 'Page': page_index + 1,
                                'X': char['origin'][0], 'Y': char['origin'][1], 'Size': size})
    return [match for prefix in prefixes for match in matches[prefix]]


def measure_images(path):
    with pymupdf.open(path) as pdf:
        return [{'Page': index + 1, 'Rect': list(item['bbox']),
            'InsidePage': pymupdf.Rect(item['bbox']) in page.rect}
            for index, page in enumerate(pdf) for item in page.get_image_info()]


def measure_note_rules(path):
    with pymupdf.open(path) as pdf:
        return [{'Page': index + 1, 'Rect': list(item['rect'])}
            for index, page in enumerate(pdf) for item in page.get_drawings()
            if item['type'] == 'f' and item['fill'] == (0., 0., 0.)
            and item['rect'].width > 50 and 0 < item['rect'].height < 2]


def rectangle_union(rectangles):
    """Envelope and covered area of emitted rectangles; clipping/ink stays separate."""
    if not rectangles:
        return [], 0.
    if any(len(rect) != 4 or any(not math.isfinite(value) for value in rect)
           or rect[2] <= rect[0] or rect[3] <= rect[1] for rect in rectangles):
        raise ValueError('Finite positive fill rectangles required')
    envelope = [min(rect[0] for rect in rectangles), min(rect[1] for rect in rectangles),
                max(rect[2] for rect in rectangles), max(rect[3] for rect in rectangles)]
    edges = sorted({value for rect in rectangles for value in (rect[1], rect[3])})
    area = 0.
    for top, bottom in zip(edges, edges[1:]):
        intervals = sorted((rect[0], rect[2]) for rect in rectangles if rect[1] < bottom and rect[3] > top)
        width, left, right = 0., None, None
        for start, end in intervals:
            if left is None:
                left, right = start, end
            elif start <= right:
                right = max(right, end)
            else:
                width += right - left
                left, right = start, end
        if left is not None:
            width += right - left
        area += width * (bottom - top)
    return envelope, area / ((envelope[2] - envelope[0]) * (envelope[3] - envelope[1]))


def measure_cell_fills(path, specifications):
    """Source-colored cell fills can be split/overpainted into several rectangles."""
    owners = {}
    for item in specifications:
        owner = {'Cell': item['Cell'], 'Side': item['Side'], 'FillHex': item['FillHex'].upper()}
        if owner['Cell'] in owners and owners[owner['Cell']] != owner:
            raise ValueError('Cell fill ownership differs between pages')
        owners[owner['Cell']] = owner
    if len({owner['FillHex'] for owner in owners.values()}) != len(owners):
        raise ValueError('Use distinct source colors to identify cell fills independently')
    result = []
    with pymupdf.open(path) as pdf:
        for page_index, page in enumerate(pdf):
            drawings = page.get_drawings()
            for cell in sorted(owners):
                owner = owners[cell]
                rgb = tuple(int(owner['FillHex'][index:index + 2], 16) for index in (0, 2, 4))
                rectangles = []
                for drawing in drawings:
                    if drawing['fill'] is None or tuple(round(value * 255) for value in drawing['fill']) != rgb:
                        continue
                    for item in drawing['items']:
                        if item[0] != 're':
                            raise ValueError('Colored fill is not an independently measured rectangle')
                        rectangles.append(list(item[1]))
                if rectangles:
                    rect, coverage = rectangle_union(rectangles)
                    result.append({'Page': page_index + 1, **owner, 'Rect': rect,
                                   'FillRectangleCount': len(rectangles), 'EnvelopeCoverage': coverage})
    return result


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
        if 'WordCellFills' in row:
            fills = measure_cell_fills(path, row['WordCellFills'])
            result['NativeCellFills'] = fills
            row_frames = []
            for page in sorted({fill['Page'] for fill in fills}):
                rect, _ = rectangle_union([fill['Rect'] for fill in fills if fill['Page'] == page])
                row_frames.append({'Page': page, 'Rect': rect})
            result['NativeRowFrames'] = row_frames
            for label, actual, expected, keys in [
                    ('CellFill', fills, row['WordCellFills'], ('Page', 'Cell')),
                    ('RowFrame', row_frames, row.get('WordRowFrames', []), ('Page',))]:
                if len(actual) != len(expected) or any(any(a[key] != e[key] for key in keys) for a, e in zip(actual, expected)):
                    result['Failures'].append(label + 'CountOrPage')
                elif expected:
                    for axis, indices in [('X', (0, 2)), ('Y', (1, 3))]:
                        error = max(abs(a['Rect'][index] - e['Rect'][index])
                                    for a, e in zip(actual, expected) for index in indices)
                        result['Max' + label + axis + 'Error'] = error
                        if axis in axes and error >= 1.:
                            result['Failures'].append(label + axis)
        if 'WordSeparatorRule' in row:
            expected_rule = row['WordSeparatorRule']
            # A report may contain filled horizontal table borders on other
            # pages. Keep every candidate on the independently pinned note page.
            rules = [rule for rule in measure_note_rules(path) if rule['Page'] == expected_rule['Page']]
            result['NativeSeparatorRules'] = rules
            if len(rules) != 1 or rules[0]['Page'] != expected_rule['Page']:
                result['Failures'].append('SeparatorCountOrPage')
            else:
                result['MaxSeparatorError'] = max(abs(a - e) for a, e in zip(rules[0]['Rect'], expected_rule['Rect']))
                if result['MaxSeparatorError'] >= 1.:
                    result['Failures'].append('SeparatorGeometry')
        if 'WordImages' in row:
            images = measure_images(path)
            expected_images = row['WordImages']
            result['NativeImages'] = images
            if len(images) != len(expected_images) or any(a['Page'] != e['Page'] for a, e in zip(images, expected_images)):
                result['Failures'].append('ImageCountOrPage')
            elif expected_images:
                for axis, indices in [('X', [0, 2]), ('Y', [1, 3])]:
                    result['MaxImage' + axis + 'Error'] = max(abs(a['Rect'][i] - e['Rect'][i])
                        for a, e in zip(images, expected_images) for i in indices)
                result['Failures'] += ['Image' + axis for axis in row.get('ImageGeometryAxes', axes)
                    if result['MaxImage' + axis + 'Error'] >= 1.]
            if any(not item['InsidePage'] for item in images):
                result['Failures'].append('OffPageImage')
        expected_prose = row.get('WordProseAnchors', [])
        prose_prefixes = [anchor['Text'] for anchor in expected_prose]
        if 'WordProseAnchorGroups' in row:
            if expected_prose:
                raise ValueError('Use one prose anchor representation')
            groups = row['WordProseAnchorGroups']
            prose_prefixes = [group['Text'] for group in groups]
            expected_prose = [{'Text': group['Text'], **anchor} for group in groups for anchor in group['Occurrences']]
        if expected_prose:
            prose = measure_prose(path, prose_prefixes)
            result['NativeProse'] = prose
            if len(prose) != len(expected_prose) or any(a['Text'] != e['Text'] or a['Page'] != e['Page'] for a, e in zip(prose, expected_prose)):
                result['Failures'].append('ProseCountOrPage')
            else:
                for axis in ['X', 'Y', 'Size']:
                    result['MaxProse' + axis + 'Error'] = max(abs(a[axis] - e[axis]) for a, e in zip(prose, expected_prose))
                result['ProseResidualAxes'] = [axis for axis in ['X', 'Y'] if result['MaxProse' + axis + 'Error'] >= 1.]
                result['Failures'] += ['Prose' + axis for axis in row.get('ProseGeometryAxes', axes)
                    if axis in result['ProseResidualAxes']]
                if result['MaxProseSizeError'] >= .1:
                    result['Failures'].append('ProsePhysicalSize')
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
        'FollowingResidualOutcomes': sum(bool(row.get('FollowingResidualAxes')) for row in rows),
        'ProseResidualOutcomes': sum(bool(row.get('ProseResidualAxes')) for row in rows), 'Rows': rows}
    args.output.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
    for row in rows:
        print(row['Id'], row['Mode'], {k: v for k, v in row.items() if k.startswith(('Max', 'Following', 'Prose')) or k == 'Failures'})
    return int(report['FailedOutcomes'] != 0)

if __name__ == '__main__':
    raise SystemExit(main())
