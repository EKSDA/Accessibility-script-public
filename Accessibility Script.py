import csv
import hashlib
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlsplit

import pycurl
import requests
from openpyxl import Workbook

BASE = Path(__file__).resolve().parent
FIRMS_LINK_LIST_URL = 'https://raw.githubusercontent.com/EKSDA/space-data-accessibility/refs/heads/main/firms_links.txt'
FIRMS_TESTS = [
    ('FIRMS_1_MODIS_C61', 'MODIS_SP', '2017-05-10', '2017-06-20', '8.3,-37.4,50.5,-0.5'),
    ('FIRMS_2_SNPP_C2', 'VIIRS_SNPP_SP', '2017-05-10', '2017-06-20', '8.3,-37.4,50.5,-0.5'),
    ('FIRMS_3_SNPP_C2', 'VIIRS_SNPP_SP', '2017-04-01', '2017-09-30', '8.3,-37.4,50.5,-0.5'),
    ('FIRMS_4_J1_C2', 'VIIRS_NOAA20_SP', '2024-01-01', '2025-01-01', 'world'),
]
TOKEN_URL = 'https://identity.dataspace.copernicus.eu/auth/realms/CDSE/protocol/openid-connect/token'
FIELDS = ['UTC', 'Iteration', 'Archive', 'Dataset', 'Request', 'Product', 'Area',
          'Start Date', 'End Date', 'Success', 'HTTP Status', 'Error',
          'Total Time (s)', 'TTFB (s)', 'Response Headers (s)', 'Download Time (s)',
          'DNS Lookup (s)', 'TCP Connect (s)', 'TLS Handshake (s)',
          'Bytes Received', 'Observed Speed (Mbps)', 'SHA256', 'Auth Time (s)', 'FIRMS List SHA256']
BATCH_FIELDS = ['Iteration', 'Archive', 'Dataset', 'Requests', 'Succeeded', 'Failed',
                'Batch Elapsed (s)', 'Successful Bytes', 'Successful Transfer Time (s)',
                'Observed Speed (Mbps)']


def lines(filename):
    return [s.strip() for s in (BASE / filename).read_text(encoding='utf-8-sig').splitlines()
            if s.strip() and not s.lstrip().startswith('#')]


def credentials(filename, prefix):
    values = lines(filename)
    keys = dict(s.split('=', 1) for s in values if s.startswith((prefix+'_USERNAME=', prefix+'_PASSWORD=')))
    if keys:
        return keys[prefix+'_USERNAME'], keys[prefix+'_PASSWORD']
    if len(values) != 2:
        raise ValueError('Credentials need a username and password on separate lines.')
    return tuple(values)


def central_firms_links():
    # Fetch once before measurement, so one run uses a single list snapshot.
    try:
        with requests.get(FIRMS_LINK_LIST_URL, timeout=(30, 60), stream=True,
                          headers={'Cache-Control': 'no-cache'}) as response:
            response.raise_for_status()
            body = bytearray()
            for chunk in response.iter_content(8192):
                body.extend(chunk)
                if len(body) > 65536:
                    raise ValueError('Central FIRMS list is too large.')
        urls = [line.strip() for line in body.decode('utf-8-sig').splitlines()
                if line.strip() and not line.lstrip().startswith('#')]
    except (requests.RequestException, UnicodeError):
        raise ValueError('Cannot retrieve the central FIRMS link list.') from None
    if len(urls) != 4 or len(set(urls)) != 4:
        raise ValueError('Central FIRMS list must contain four distinct download URLs.')
    for url in urls:
        parsed = urlsplit(url)
        if (parsed.scheme != 'https' or not parsed.hostname
                or not parsed.hostname.endswith('.modaps.eosdis.nasa.gov')
                or parsed.username or parsed.password):
            raise ValueError('Central FIRMS list must contain HTTPS FIRMS download URLs.')
    return urls, hashlib.sha256(body).hexdigest()


def valid_body(prefix, content_type, firms=False):
    stripped = prefix.lstrip().lower()
    if not prefix or 'html' in content_type.lower() or stripped.startswith((b'<', b'{', b'[')):
        return False
    if firms:
        # Check ZIP signature; full archive integrity is not checked because
        # the benchmark discards file bytes after counting and hashing them.
        return prefix.startswith((b'PK\x03\x04', b'PK\x05\x06'))
    return True


def fetch_curl(url, token=None, firms=False):
    """Keep bounded sample only; retain partial byte counts on failure."""
    row = {'Success': False, 'Bytes Received': 0, 'Error': ''}
    sample, digest = bytearray(), hashlib.sha256()
    def receive(chunk):
        row['Bytes Received'] += len(chunk)
        digest.update(chunk)
        if len(sample) < 4096:
            sample.extend(chunk[:4096-len(sample)])
        return len(chunk)
    curl = pycurl.Curl()
    try:
        curl.setopt(pycurl.URL, url)
        curl.setopt(pycurl.FOLLOWLOCATION, True)
        curl.setopt(pycurl.MAXREDIRS, 10)
        curl.setopt(pycurl.CONNECTTIMEOUT, 60)
        curl.setopt(pycurl.TIMEOUT, 1200)
        curl.setopt(pycurl.WRITEFUNCTION, receive)
        curl.setopt(pycurl.FAILONERROR, True)
        # Never enable UNRESTRICTED_AUTH: do not forward tokens to other hosts.
        if token:
            curl.setopt(pycurl.HTTPHEADER, ['Authorization: Bearer '+token])
        curl.perform()
        code = curl.getinfo(pycurl.RESPONSE_CODE)
        row['Success'] = code == 200 and valid_body(sample, curl.getinfo(pycurl.CONTENT_TYPE) or '', firms)
        if not row['Success']:
            row['Error'] = 'Unexpected status or non-data response'
    except pycurl.error as exc:
        row['Error'] = 'Curl error code '+str(exc.args[0])
    finally:
        dns = curl.getinfo(pycurl.NAMELOOKUP_TIME)
        tcp = curl.getinfo(pycurl.CONNECT_TIME)
        tls = curl.getinfo(pycurl.APPCONNECT_TIME)
        ttfb = curl.getinfo(pycurl.STARTTRANSFER_TIME)
        total = curl.getinfo(pycurl.TOTAL_TIME)
        row.update({'HTTP Status': curl.getinfo(pycurl.RESPONSE_CODE),
                    'Total Time (s)': total, 'TTFB (s)': ttfb,
                    'Download Time (s)': max(0, total-ttfb),
                    'DNS Lookup (s)': dns, 'TCP Connect (s)': max(0, tcp-dns),
                    'TLS Handshake (s)': max(0, tls-tcp) if tls else None})
        curl.close()
    finish_metrics(row, digest)
    return row


def finish_metrics(row, digest):
    if row['Success']:
        row['SHA256'] = digest.hexdigest()
        duration = row.get('Download Time (s)', 0)
        row['Observed Speed (Mbps)'] = row['Bytes Received']*8/duration/1_000_000 if duration else None


class EarthdataSession(requests.Session):
    def rebuild_auth(self, prepared, response):
        old, new = urlsplit(response.request.url), urlsplit(prepared.url)
        if new.scheme != 'https' or (old.hostname != new.hostname
                and new.hostname != 'urs.earthdata.nasa.gov'
                and old.hostname != 'urs.earthdata.nasa.gov'):
            prepared.headers.pop('Authorization', None)


def fetch_earthdata(url, session):
    row = {'Success': False, 'Bytes Received': 0, 'Error': ''}
    start, digest, sample = time.perf_counter(), hashlib.sha256(), bytearray()
    try:
        with session.get(url, stream=True, timeout=(60, 1200)) as response:
            row['Response Headers (s)'] = time.perf_counter()-start
            row['HTTP Status'] = response.status_code
            response.raise_for_status()
            if response.status_code != 200:
                row['Error'] = 'Unexpected HTTP status'
                return row
            for chunk in response.iter_content(1024*1024):
                row['Bytes Received'] += len(chunk)
                digest.update(chunk)
                if len(sample) < 4096:
                    sample.extend(chunk[:4096-len(sample)])
            row['Success'] = valid_body(sample, response.headers.get('Content-Type', ''))
            if not row['Success']:
                row['Error'] = 'Empty, login or non-data response'
    except requests.RequestException as exc:
        row['Error'] = type(exc).__name__  # Exception text can contain secret URLs.
    finally:
        row['Total Time (s)'] = time.perf_counter()-start
        if 'Response Headers (s)' in row:
            row['Download Time (s)'] = max(0, row['Total Time (s)']-row['Response Headers (s)'])
    finish_metrics(row, digest)
    return row


def cdse_token(creds):
    with requests.post(TOKEN_URL, data={'client_id': 'cdse-public', 'grant_type': 'password',
                       'username': creds[0], 'password': creds[1]}, timeout=(60, 120)) as response:
        response.raise_for_status()
        return response.json()['access_token']


def save_excel(path, rows, batches):
    book = Workbook()
    for sheet, fields, data in [(book.active, FIELDS, rows),
                                (book.create_sheet('Batches'), BATCH_FIELDS, batches)]:
        if sheet == book.active:
            sheet.title = 'Requests'
        sheet.append(fields)
        for row in data:
            sheet.append([row.get(field) for field in fields])
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
    book.save(path)


def main():
    print('Fetching the current central FIRMS link list...')
    firms_urls, list_hash = central_firms_links()
    earth_creds = credentials('earthdata_credentials.txt', 'EARTHDATA')
    cdse_creds = credentials('cdse_credentials.txt', 'CDSE')
    tests = []
    for url, (label, product, start, end, area) in zip(firms_urls, FIRMS_TESTS):
        settings = {'Product': product, 'Area': area, 'Start Date': start,
                    'End Date': end, 'FIRMS List SHA256': list_hash}
        tests.append(('FIRMS', label, [(url, settings)]))
    for filename in ('earthdata_links.txt', 'earthdata_links_1.txt'):
        tests.append(('EARTHDATA', Path(filename).stem, [(url, {}) for url in lines(filename)]))
    for number, url in enumerate(lines('cdse_links.txt'), 1):
        tests.append(('CDSE', f'CDSE_{number}', [(url, {})]))
    if any(not requests_for_test for _, _, requests_for_test in tests):
        raise ValueError('A link file is empty.')
    if any(urlsplit(url).scheme != 'https' for _, _, items in tests for url, _ in items):
        raise ValueError('Link files must contain HTTPS download URLs.')
    iterations = int(input('Number of iterations: '))
    if iterations < 1:
        raise ValueError('Iterations must be at least 1.')
    stamp = datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S_%f')
    output = BASE / ('accessibility_results_'+stamp)
    rows, batches = [], []
    print('FIRMS: four complete ZIP downloads per iteration. Central list loaded.')
    with output.with_suffix('.csv').open('w', newline='', encoding='utf-8') as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS)
        writer.writeheader()
        stream.flush()
        try:
            for iteration in range(1, iterations+1):
                for archive, label, items in tests:
                    batch_start = time.perf_counter()
                    current = []
                    with EarthdataSession() as earth:
                        earth.auth = earth_creds
                        for number, (url, settings) in enumerate(items, 1):
                            print(f'Iteration {iteration}/{iterations}: {label} {number}/{len(items)}', flush=True)
                            auth_time = 0
                            utc = datetime.now(timezone.utc).isoformat()
                            if archive == 'FIRMS':
                                result = fetch_curl(url, firms=True)
                            elif archive == 'EARTHDATA':
                                result = fetch_earthdata(url, earth)
                            else:
                                auth_start = time.perf_counter()
                                try:
                                    token = cdse_token(cdse_creds)
                                except (requests.RequestException, ValueError, KeyError) as exc:
                                    result = {'Success': False, 'Bytes Received': 0,
                                              'Error': 'CDSE authentication: '+type(exc).__name__}
                                else:
                                    auth_time = time.perf_counter()-auth_start
                                    result = fetch_curl(url, token=token)
                                if not auth_time:
                                    auth_time = time.perf_counter()-auth_start
                            row = {'UTC': utc, 'Iteration': iteration, 'Archive': archive,
                                   'Dataset': label, 'Request': number, 'Auth Time (s)': auth_time,
                                   **settings, **result}
                            rows.append(row)
                            current.append(row)
                            writer.writerow(row)
                            stream.flush()
                            print('  OK' if row['Success'] else '  FAILED: '+row['Error'])
                    good = [row for row in current if row['Success']]
                    size = sum(row['Bytes Received'] for row in good)
                    duration = sum(row.get('Download Time (s)', 0) for row in good)
                    batches.append({'Iteration': iteration, 'Archive': archive, 'Dataset': label,
                                    'Requests': len(items), 'Succeeded': len(good), 'Failed': len(items)-len(good),
                                    'Batch Elapsed (s)': time.perf_counter()-batch_start,
                                    'Successful Bytes': size, 'Successful Transfer Time (s)': duration,
                                    'Observed Speed (Mbps)': size*8/duration/1_000_000 if duration else None})
        except KeyboardInterrupt:
            print('\nStopped. Completed request results have been kept.')
        finally:
            save_excel(output.with_suffix('.xlsx'), rows, batches)
            print('Results:', output.with_suffix('.csv').name, 'and', output.with_suffix('.xlsx').name)


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, KeyError, IndexError) as exc:
        print('Could not run: check the input files and iteration number ('+type(exc).__name__+').')
        raise SystemExit(1)
