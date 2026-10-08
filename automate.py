"""Generate and optionally email certificates without a web application."""
import argparse
import json
import os
import re
import sys
import time
from pathlib import Path
from dotenv import dotenv_values
from certificates import certificate_id, connect_log, overlay_name, read_participants, send_certificate, SMTPBatch


def smtp_settings(config, base):
    # Read per run without leaking one configuration's secrets into another.
    values = {**dotenv_values(base / '.env', encoding='utf-8-sig', interpolate=False), **os.environ}
    smtp = dict(config.get('smtp', {}))
    for field in ('host', 'sender', 'username', 'password'):
        key = 'SMTP_' + field.upper()
        if key in values:
            smtp[field] = values[key]
    if 'SMTP_PORT' in values:
        try:
            smtp['port'] = int(values['SMTP_PORT'])
        except (TypeError, ValueError):
            raise ValueError('SMTP_PORT must be an integer from 1 to 65535.') from None
    if 'SMTP_SSL' in values:
        setting = (values['SMTP_SSL'] or '').strip().lower()
        if setting not in ('true', 'false'):
            raise ValueError('SMTP_SSL must be true or false.')
        smtp['ssl'] = setting == 'true'
    if not smtp.get('username'):
        smtp['username'] = smtp.get('sender', '')
    return smtp


def run(config_path, send=False, resend=False):
    if resend and not send:
        raise ValueError('--resend requires --send.')
    started = time.perf_counter()
    config_path = Path(config_path).resolve()
    config = json.loads(config_path.read_text(encoding='utf-8'))
    base = config_path.parent
    for field in ('event_id', 'organizer'):
        if not isinstance(config.get(field), str) or not config[field].strip():
            raise ValueError(f'Enter {field} in the configuration.')
    placement = config['name']
    if not (0 < placement['x_percent'] < 100 and 0 < placement['y_percent'] < 100 and 8 <= placement['font_size'] <= 200):
        raise ValueError('Invalid name position or font size (8–200 points).')
    if not re.fullmatch(r'#[0-9a-fA-F]{6}', placement['color']):
        raise ValueError('Name color must use #RRGGBB format.')
    people, errors = read_participants(base / config['excel'], config.get('sheet'))
    if errors:
        raise ValueError('\n'.join(f"Excel row {e['row']}: {e['problem']}" for e in errors))
    if not people:
        raise ValueError('The Excel sheet contains no participants.')
    if send:
        smtp = smtp_settings(config, base)
        if not all(smtp.get(k) for k in ('host', 'port', 'sender', 'username', 'password')):
            raise ValueError('Complete SMTP settings in .env beside config.json (see .env.example). Set SMTP_SENDER and SMTP_PASSWORD; SMTP_USERNAME defaults to the sender.')
        if not isinstance(smtp.get('ssl'), bool) or not isinstance(smtp['port'], int) or not 1 <= smtp['port'] <= 65535:
            raise ValueError('SMTP ssl must be true/false and port must be 1–65535.')
        for field in ('sender', 'username', 'host'):
            if any(ord(c) < 32 for c in smtp[field]):
                raise ValueError(f'Invalid SMTP {field}.')
    template_path = base / config['template']
    template = template_path.read_bytes()
    # Validate and render the full batch before sending any messages.
    rendered = []
    for index, person in enumerate(people, 1):
        rendered.append(overlay_name(template, template_path.name, person['name'], placement['x_percent'],
                                     placement['y_percent'], placement['font_size'], placement['color']))
        if index % 25 == 0 or index == len(people):
            print(f'Prepared {index}/{len(people)} certificates', flush=True)
    output = base / config['output']
    output.mkdir(parents=True, exist_ok=True)
    for person, pdf in zip(people, rendered):
        key = certificate_id(config['event_id'], person['email'])
        (output / f'Grainathon-{key}.pdf').write_bytes(pdf)
    print(f'Generated {len(people)} certificates in {output} ({time.perf_counter()-started:.1f}s)', flush=True)
    if not send:
        print('Generation only. Add --send to email the certificates.')
        return 0
    db = connect_log(base / config['delivery_log'])
    transport = SMTPBatch(smtp)
    try:
        for person, pdf in zip(people, rendered):
            status = send_certificate(person, config, pdf, smtp, db, transport=transport, resend=resend)
            print(f"{person['email']}: {status}", flush=True)
            key = certificate_id(config['event_id'], person['email'])
            recorded = db.execute('SELECT status FROM deliveries WHERE id=?', (key,)).fetchone()
            if recorded and recorded[0] != 'accepted by SMTP':
                print('Stopped: check the delivery log and mail provider before retrying.', file=sys.stderr)
                return 1
    finally:
        transport.close()
        db.close()
    print(f'Batch completed in {time.perf_counter()-started:.1f}s', flush=True)
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', default=str(Path(__file__).with_name('config.json')))
    parser.add_argument('--send', action='store_true', help='Email the generated certificates.')
    parser.add_argument('--resend', action='store_true', help='With --send, send again to previously accepted recipients in this Excel sheet.')
    args = parser.parse_args()
    try:
        return run(args.config, args.send, args.resend)
    except (OSError, ValueError, KeyError, TypeError) as exc:
        print(f'Error: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
