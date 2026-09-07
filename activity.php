<?php
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store, max-age=0');
header('Pragma: no-cache');
header('X-Content-Type-Options: nosniff');
header('Referrer-Policy: same-origin');

const ACTIVE_WINDOW_SECONDS = 300;
const VISIT_WINDOW_SECONDS = 1800;
const SESSION_RETENTION_SECONDS = 2764800; // 32 days
const RECENT_RETENTION_SECONDS = 7776000;  // 90 days
const MAX_RECENT_EVENTS = 40;

function jsonResponse(array $payload, int $status = 200): never
{
    http_response_code($status);
    echo json_encode($payload, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
    exit;
}

function cleanCountry(mixed $value): string
{
    $code = strtoupper(trim((string)$value));
    return preg_match('/^[A-Z]{2}$/', $code) && !in_array($code, ['XX', 'T1'], true)
        ? $code
        : 'UN';
}

function emptyMonth(): array
{
    return [
        'visits' => 0,
        'docxSelected' => 0,
        'reviewEdits' => 0,
        'exports' => 0,
        'countries' => [],
    ];
}

function emptyData(): array
{
    return [
        'version' => 1,
        'updated' => gmdate('c'),
        'months' => [],
        'sessions' => [],
        'recent' => [],
    ];
}

function publicSnapshot(array $data, int $now): array
{
    $activeNow = 0;
    foreach ($data['sessions'] as $session) {
        if (($session['lastSeen'] ?? 0) >= $now - ACTIVE_WINDOW_SECONDS) {
            $activeNow++;
        }
    }

    $months = $data['months'];
    $currentMonth = gmdate('Y-m', $now);
    if (!isset($months[$currentMonth])) {
        $months[$currentMonth] = emptyMonth();
    }
    krsort($months);
    $months = array_slice($months, 0, 18, true);

    return [
        'ok' => true,
        'activeNow' => $activeNow,
        'activeWindowMinutes' => (int)(ACTIVE_WINDOW_SECONDS / 60),
        'updated' => $data['updated'],
        'months' => $months,
        'recent' => array_values(array_slice($data['recent'], 0, 16)),
        'privacy' => 'Anonymous aggregate events only. No names, filenames, document contents, or IP addresses are stored.',
    ];
}

$method = $_SERVER['REQUEST_METHOD'] ?? 'GET';
if (!in_array($method, ['GET', 'POST'], true)) {
    header('Allow: GET, POST');
    jsonResponse(['ok' => false, 'error' => 'Method not allowed'], 405);
}

if ($method === 'POST') {
    $origin = $_SERVER['HTTP_ORIGIN'] ?? '';
    $host = $_SERVER['HTTP_HOST'] ?? '';
    if ($origin !== '' && strcasecmp((string)parse_url($origin, PHP_URL_HOST), preg_replace('/:\d+$/', '', $host)) !== 0) {
        jsonResponse(['ok' => false, 'error' => 'Cross-origin request rejected'], 403);
    }
    if ((int)($_SERVER['CONTENT_LENGTH'] ?? 0) > 2048) {
        jsonResponse(['ok' => false, 'error' => 'Request too large'], 413);
    }
}

$storageDir = getenv('QTI_ACTIVITY_DIR') ?: dirname(__DIR__, 3) . DIRECTORY_SEPARATOR . 'qti_activity';
if (!is_dir($storageDir) && !mkdir($storageDir, 0700, true) && !is_dir($storageDir)) {
    jsonResponse(['ok' => false, 'error' => 'Analytics storage is unavailable'], 503);
}
$storageFile = $storageDir . DIRECTORY_SEPARATOR . 'docx_to_qti_activity.json';
$handle = fopen($storageFile, 'c+');
if ($handle === false || !flock($handle, LOCK_EX)) {
    jsonResponse(['ok' => false, 'error' => 'Analytics storage is busy'], 503);
}

$now = time();
$modified = false;

try {
    rewind($handle);
    $raw = stream_get_contents($handle);
    $data = $raw !== false && trim($raw) !== '' ? json_decode($raw, true) : null;
    if (!is_array($data) || !isset($data['months'], $data['sessions'], $data['recent'])) {
        $data = emptyData();
        $modified = true;
    }

    foreach ($data['sessions'] as $key => $session) {
        if (($session['lastSeen'] ?? 0) < $now - SESSION_RETENTION_SECONDS) {
            unset($data['sessions'][$key]);
            $modified = true;
        }
    }
    $filteredRecent = array_values(array_filter(
        $data['recent'],
        static fn(array $entry): bool => ($entry['time'] ?? 0) >= $now - RECENT_RETENTION_SECONDS
    ));
    if (count($filteredRecent) !== count($data['recent'])) {
        $data['recent'] = $filteredRecent;
        $modified = true;
    }

    if ($method === 'POST') {
        $body = json_decode((string)file_get_contents('php://input'), true);
        if (!is_array($body)) {
            jsonResponse(['ok' => false, 'error' => 'Invalid JSON body'], 400);
        }

        $event = (string)($body['event'] ?? '');
        $allowedEvents = ['visit', 'heartbeat', 'docx_selected', 'review_edit', 'export'];
        if (!in_array($event, $allowedEvents, true)) {
            jsonResponse(['ok' => false, 'error' => 'Unsupported event'], 422);
        }

        $sessionId = trim((string)($body['session'] ?? ''));
        if (!preg_match('/^[A-Za-z0-9-]{16,80}$/', $sessionId)) {
            jsonResponse(['ok' => false, 'error' => 'Invalid anonymous session'], 422);
        }

        $cfCountry = cleanCountry($_SERVER['HTTP_CF_IPCOUNTRY'] ?? '');
        $country = $cfCountry !== 'UN' ? $cfCountry : cleanCountry($body['country'] ?? '');
        $sessionKey = hash('sha256', $sessionId);
        $session = $data['sessions'][$sessionKey] ?? [
            'lastSeen' => 0,
            'lastVisit' => 0,
            'lastEvents' => [],
            'country' => $country,
        ];
        $session['lastSeen'] = $now;
        if ($country !== 'UN') {
            $session['country'] = $country;
        } else {
            $country = $session['country'] ?? 'UN';
        }

        $monthKey = gmdate('Y-m', $now);
        if (!isset($data['months'][$monthKey])) {
            $data['months'][$monthKey] = emptyMonth();
        }
        $month = &$data['months'][$monthKey];
        $recordRecent = false;

        if ($event === 'visit') {
            if (($session['lastVisit'] ?? 0) < $now - VISIT_WINDOW_SECONDS) {
                $session['lastVisit'] = $now;
                $month['visits']++;
                $month['countries'][$country] = ($month['countries'][$country] ?? 0) + 1;
                $recordRecent = true;
            }
        } elseif ($event !== 'heartbeat') {
            $minimumGap = $event === 'review_edit' ? 20 : 4;
            $lastEvent = (int)($session['lastEvents'][$event] ?? 0);
            if ($lastEvent <= $now - $minimumGap) {
                $field = [
                    'docx_selected' => 'docxSelected',
                    'review_edit' => 'reviewEdits',
                    'export' => 'exports',
                ][$event];
                $month[$field]++;
                $session['lastEvents'][$event] = $now;
                $recordRecent = true;
            }
        }

        if ($recordRecent) {
            array_unshift($data['recent'], [
                'time' => $now,
                'event' => $event,
                'country' => $country,
            ]);
            $data['recent'] = array_slice($data['recent'], 0, MAX_RECENT_EVENTS);
        }

        $data['sessions'][$sessionKey] = $session;
        $data['updated'] = gmdate('c', $now);
        $modified = true;
    }

    if ($modified) {
        rewind($handle);
        ftruncate($handle, 0);
        fwrite($handle, json_encode($data, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
        fflush($handle);
    }

    $snapshot = publicSnapshot($data, $now);
} finally {
    flock($handle, LOCK_UN);
    fclose($handle);
}

jsonResponse($snapshot);
