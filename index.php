<?php
declare(strict_types=1);

require_once __DIR__ . '/config.php';

mysqli_set_charset($conn, 'utf8mb4');
date_default_timezone_set('Asia/Jakarta');
session_start();

const MAX_UPLOAD_BYTES = 10485760;
const PAGE_SIZE = 25;

function showDatabaseUnavailable(): never
{
    http_response_code(503);
    echo '<!doctype html><html lang="id"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">';
    echo '<title>Database belum tersedia</title><link rel="stylesheet" href="style.css"><main class="page-shell database-error">';
    echo '<section class="panel database-error-panel"><p class="eyebrow">KONEKSI DIPERLUKAN</p>';
    echo '<h1>Database belum dapat diakses</h1><p class="welcome-copy">Jalankan layanan MySQL melalui XAMPP dan pastikan database <strong>data_sblw</strong> serta pengaturan di <strong>config.php</strong> sudah benar.</p></section></main></html>';
    exit;
}

try {
    require_once __DIR__ . '/config.php';
    mysqli_set_charset($conn, 'utf8mb4');
    date_default_timezone_set('Asia/Jakarta');
    if (session_status() !== PHP_SESSION_ACTIVE) {
        session_start();
    }

    $conn->query(
        "CREATE TABLE IF NOT EXISTS billing_customers (
        idpel VARCHAR(32) NOT NULL PRIMARY KEY,
        koded VARCHAR(64) NOT NULL DEFAULT '',
        nama VARCHAR(160) NOT NULL,
        alamat TEXT NOT NULL,
        tarif VARCHAR(32) NOT NULL DEFAULT '',
        daya BIGINT UNSIGNED NOT NULL DEFAULT 0,
        rptag BIGINT UNSIGNED NOT NULL DEFAULT 0,
        rpbk BIGINT UNSIGNED NOT NULL DEFAULT 0,
        total BIGINT UNSIGNED NOT NULL DEFAULT 0,
        kontak_pelanggan VARCHAR(40) NOT NULL DEFAULT '',
        no_wa VARCHAR(40) NOT NULL DEFAULT '',
        payment_status VARCHAR(16) NOT NULL DEFAULT 'unknown',
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_billing_status (payment_status),
        INDEX idx_billing_name (nama)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci"
    );
    $conn->query(
        "CREATE TABLE IF NOT EXISTS billing_app_meta (
        meta_key VARCHAR(40) NOT NULL PRIMARY KEY,
        meta_value VARCHAR(64) NOT NULL DEFAULT ''
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci"
    );
    $conn->query(
        "CREATE TABLE IF NOT EXISTS billing_whatsapp_clicks (
        idpel VARCHAR(32) NOT NULL PRIMARY KEY,
        clicked_on DATE NOT NULL
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci"
    );
} catch (mysqli_sql_exception $exception) {
    error_log('Billing database unavailable: ' . $exception->getMessage());
    showDatabaseUnavailable();
}

function h(string $value): string
{
    return htmlspecialchars($value, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

function setFlash(string $type, string $message): void
{
    $_SESSION['billing_flash'] = ['type' => $type, 'message' => $message];
}

function redirectToPage(): never
{
    header('Location: index.php');
    exit;
}

function getMeta(mysqli $conn, string $key): string
{
    $stmt = $conn->prepare('SELECT meta_value FROM billing_app_meta WHERE meta_key = ?');
    $stmt->bind_param('s', $key);
    $stmt->execute();
    $stmt->bind_result($value);
    $stmt->fetch();
    $stmt->close();
    return $value ?? '';
}

function setMeta(mysqli $conn, string $key, string $value): void
{
    $stmt = $conn->prepare(
        'INSERT INTO billing_app_meta (meta_key, meta_value) VALUES (?, ?)
         ON DUPLICATE KEY UPDATE meta_value = VALUES(meta_value)'
    );
    $stmt->bind_param('ss', $key, $value);
    $stmt->execute();
    $stmt->close();
}

function normalizeHeader(string $header): string
{
    $header = preg_replace('/^\xEF\xBB\xBF/', '', $header) ?? $header;
    return strtoupper((string) preg_replace('/[^A-Z0-9]/', '', strtoupper(trim($header))));
}

function getCellText(DOMXPath $xpath, DOMElement $cell, string $path): string
{
    $node = $xpath->query($path, $cell)?->item(0);
    return $node?->textContent ?? '';
}

function readXlsxRows(string $path): array
{
    if (!class_exists(ZipArchive::class)) {
        throw new RuntimeException('Ekstensi PHP ZipArchive belum aktif. Aktifkan ekstensi zip di PHP.');
    }

    $zip = new ZipArchive();
    if ($zip->open($path) !== true) {
        throw new RuntimeException('File Excel tidak dapat dibuka. Pastikan file berformat .xlsx yang valid.');
    }

    try {
        $sharedStrings = [];
        $sharedXml = $zip->getFromName('xl/sharedStrings.xml');
        if ($sharedXml !== false) {
            $sharedDoc = new DOMDocument();
            if (!$sharedDoc->loadXML($sharedXml, LIBXML_NONET)) {
                throw new RuntimeException('Isi shared strings pada file Excel tidak valid.');
            }
            $sharedXPath = new DOMXPath($sharedDoc);
            foreach ($sharedXPath->query('//*[local-name()="si"]') as $item) {
                $parts = [];
                foreach ($sharedXPath->query('.//*[local-name()="t"]', $item) as $textNode) {
                    $parts[] = $textNode->textContent;
                }
                $sharedStrings[] = implode('', $parts);
            }
        }

        $sheetName = 'xl/worksheets/sheet1.xml';
        $workbookXml = $zip->getFromName('xl/workbook.xml');
        $relationsXml = $zip->getFromName('xl/_rels/workbook.xml.rels');
        if ($workbookXml !== false && $relationsXml !== false) {
            $workbook = new DOMDocument();
            $relations = new DOMDocument();
            if (!$workbook->loadXML($workbookXml, LIBXML_NONET) || !$relations->loadXML($relationsXml, LIBXML_NONET)) {
                throw new RuntimeException('Struktur workbook Excel tidak valid.');
            }
            $workbookXPath = new DOMXPath($workbook);
            $relationsXPath = new DOMXPath($relations);
            $firstSheet = $workbookXPath->query('//*[local-name()="sheets"]/*[local-name()="sheet"]')->item(0);
            if ($firstSheet instanceof DOMElement) {
                $relationId = $firstSheet->getAttributeNS(
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                    'id'
                );
                foreach ($relationsXPath->query('//*[local-name()="Relationship"]') as $relation) {
                    if ($relation instanceof DOMElement && $relation->getAttribute('Id') === $relationId) {
                        $target = str_replace('\\', '/', $relation->getAttribute('Target'));
                        $sheetName = ltrim(str_starts_with($target, '/') ? $target : 'xl/' . $target, '/');
                        break;
                    }
                }
            }
        }

        $sheetXml = $zip->getFromName($sheetName);
        if ($sheetXml === false) {
            throw new RuntimeException('Sheet pertama tidak ditemukan di dalam file Excel.');
        }
        $sheet = new DOMDocument();
        if (!$sheet->loadXML($sheetXml, LIBXML_NONET)) {
            throw new RuntimeException('Data sheet pada file Excel tidak valid.');
        }
        $xpath = new DOMXPath($sheet);
        $rows = [];
        foreach ($xpath->query('//*[local-name()="sheetData"]/*[local-name()="row"]') as $rowNode) {
            $row = [];
            foreach ($xpath->query('./*[local-name()="c"]', $rowNode) as $cell) {
                if (!$cell instanceof DOMElement || !preg_match('/^([A-Z]+)/', $cell->getAttribute('r'), $match)) {
                    continue;
                }
                $column = 0;
                foreach (str_split($match[1]) as $letter) {
                    $column = ($column * 26) + ord($letter) - 64;
                }
                $column--;
                $type = $cell->getAttribute('t');
                if ($type === 'inlineStr') {
                    $value = getCellText($xpath, $cell, './/*[local-name()="t"]');
                } else {
                    $value = getCellText($xpath, $cell, './*[local-name()="v"]');
                    if ($type === 's' && $value !== '') {
                        $value = $sharedStrings[(int) $value] ?? '';
                    }
                }
                $row[$column] = trim($value);
            }
            if ($row !== []) {
                ksort($row);
                $rows[] = $row;
            }
        }
        return $rows;
    } finally {
        $zip->close();
    }
}

function readCsvRows(string $path): array
{
    $handle = fopen($path, 'rb');
    if ($handle === false) {
        throw new RuntimeException('File CSV tidak dapat dibaca.');
    }

    try {
        $firstLine = fgets($handle);
        if ($firstLine === false) {
            throw new RuntimeException('File CSV kosong.');
        }
        $delimiters = [',', ';', "\t"];
        $delimiter = ',';
        $maxColumns = 0;
        foreach ($delimiters as $candidate) {
            $count = count(str_getcsv($firstLine, $candidate, '"', '\\'));
            if ($count > $maxColumns) {
                $maxColumns = $count;
                $delimiter = $candidate;
            }
        }
        rewind($handle);
        $rows = [];
        while (($row = fgetcsv($handle, 0, $delimiter, '"', '\\')) !== false) {
            if (count($row) === 1 && trim((string) $row[0]) === '') {
                continue;
            }
            $rows[] = array_map(static fn($value): string => trim((string) $value), $row);
        }
        return $rows;
    } finally {
        fclose($handle);
    }
}

function readTabularRows(string $path, string $extension): array
{
    if ($extension === 'xlsx') {
        return readXlsxRows($path);
    }
    if ($extension === 'csv') {
        return readCsvRows($path);
    }
    throw new RuntimeException('Format file tidak didukung. Gunakan .xlsx atau .csv.');
}

function makeHeaderMap(array $headerRow): array
{
    $map = [];
    foreach ($headerRow as $index => $header) {
        $key = normalizeHeader((string) $header);
        if ($key !== '') {
            if ($key === 'NOIDPEL') {
                $key = 'IDPEL';
            }
            $map[$key] = $index;
        }
    }
    return $map;
}

function bindStatementParams(mysqli_stmt $stmt, string $types, array &$params): void
{
    $bindings = [$types];
    foreach ($params as $index => &$value) {
        $bindings[] = &$value;
    }
    unset($value);
    call_user_func_array([$stmt, 'bind_param'], $bindings);
}

function rowValue(array $row, array $headerMap, string $header): string
{
    $index = $headerMap[$header] ?? null;
    return $index !== null ? trim((string) ($row[$index] ?? '')) : '';
}

function parseAmount(string $value, string $field, int $rowNumber): int
{
    $value = trim($value);
    if ($value === '') {
        return 0;
    }
    $digits = preg_replace('/\D/', '', $value);
    if ($digits === '' || strlen($digits) > 15) {
        throw new RuntimeException("Nilai {$field} pada baris {$rowNumber} tidak valid.");
    }
    return (int) $digits;
}

function formatRupiah(int|string $amount): string
{
    return 'Rp ' . number_format((int) $amount, 0, ',', '.');
}

function normalizeWhatsapp(string $phone): string
{
    $digits = preg_replace('/\D/', '', $phone) ?? '';
    if (str_starts_with($digits, '0')) {
        $digits = '62' . substr($digits, 1);
    } elseif (str_starts_with($digits, '8')) {
        $digits = '62' . $digits;
    }
    return $digits;
}

function respondJson(int $status, array $payload): never
{
    http_response_code($status);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($payload);
    exit;
}

if (empty($_SESSION['billing_csrf'])) {
    $_SESSION['billing_csrf'] = bin2hex(random_bytes(32));
}

if (($_SERVER['REQUEST_METHOD'] ?? 'GET') === 'POST') {
    $action = (string) ($_POST['action'] ?? '');
    $submittedToken = (string) ($_POST['csrf_token'] ?? '');
    if (!hash_equals($_SESSION['billing_csrf'], $submittedToken)) {
        if ($action === 'mark_whatsapp') {
            respondJson(403, ['success' => false, 'message' => 'Sesi keamanan tidak valid. Muat ulang halaman.']);
        }
        setFlash('error', 'Sesi keamanan tidak valid. Muat ulang halaman, lalu coba kembali.');
        redirectToPage();
    }

    if ($action === 'mark_whatsapp') {
        $idpel = trim((string) ($_POST['idpel'] ?? ''));
        if ($idpel === '' || strlen($idpel) > 32) {
            respondJson(400, ['success' => false, 'message' => 'IDPEL tidak valid.']);
        }

        try {
            $customerStmt = $conn->prepare(
                "SELECT 1 FROM billing_customers WHERE idpel = ? AND payment_status = 'unpaid' LIMIT 1"
            );
            $customerStmt->bind_param('s', $idpel);
            $customerStmt->execute();
            $customerStmt->store_result();
            $isUnpaidCustomer = $customerStmt->num_rows > 0;
            $customerStmt->close();
            if (!$isUnpaidCustomer) {
                respondJson(404, ['success' => false, 'message' => 'Pelanggan belum lunas tidak ditemukan.']);
            }

            $clickedOn = date('Y-m-d');
            $clickStmt = $conn->prepare(
                'INSERT INTO billing_whatsapp_clicks (idpel, clicked_on) VALUES (?, ?)
                 ON DUPLICATE KEY UPDATE clicked_on = VALUES(clicked_on)'
            );
            $clickStmt->bind_param('ss', $idpel, $clickedOn);
            $clickStmt->execute();
            $clickStmt->close();
            respondJson(200, ['success' => true]);
        } catch (Throwable $exception) {
            error_log('Unable to record WhatsApp click: ' . $exception->getMessage());
            respondJson(500, ['success' => false, 'message' => 'Penanda WhatsApp tidak dapat disimpan.']);
        }
    }

    $file = $_FILES['spreadsheet'] ?? null;

    try {
        if (!is_array($file) || ($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
            throw new RuntimeException('Pilih file yang valid sebelum mengunggah.');
        }
        if ((int) $file['size'] <= 0 || (int) $file['size'] > MAX_UPLOAD_BYTES) {
            throw new RuntimeException('Ukuran file harus lebih dari 0 dan maksimal 10 MB.');
        }
        $extension = strtolower(pathinfo((string) $file['name'], PATHINFO_EXTENSION));
        if (!in_array($extension, ['xlsx', 'csv'], true)) {
            throw new RuntimeException('Gunakan file Excel .xlsx atau file .csv.');
        }

        if ($action === 'import_master') {
            if ($extension !== 'xlsx') {
                throw new RuntimeException('Data master bulanan harus menggunakan file Excel .xlsx.');
            }
            $rows = readTabularRows((string) $file['tmp_name'], $extension);
            if (count($rows) < 2) {
                throw new RuntimeException('File master harus memiliki baris header dan setidaknya satu data pelanggan.');
            }
            $headers = makeHeaderMap($rows[0]);
            foreach (['IDPEL', 'NAMA'] as $required) {
                if (!array_key_exists($required, $headers)) {
                    throw new RuntimeException("Kolom {$required} tidak ditemukan pada file master.");
                }
            }

            $customers = [];
            foreach (array_slice($rows, 1) as $offset => $row) {
                $rowNumber = $offset + 2;
                $idpel = preg_replace('/\s+/', '', rowValue($row, $headers, 'IDPEL')) ?? '';
                if ($idpel === '') {
                    continue;
                }
                $nama = rowValue($row, $headers, 'NAMA');
                if ($nama === '') {
                    throw new RuntimeException("Nama pelanggan kosong pada baris {$rowNumber}.");
                }
                if (isset($customers[$idpel])) {
                    throw new RuntimeException("IDPEL duplikat {$idpel} ditemukan pada baris {$rowNumber}.");
                }
                $rptag = parseAmount(rowValue($row, $headers, 'RPTAG'), 'RPTAG', $rowNumber);
                $rpbk = parseAmount(rowValue($row, $headers, 'RPBK'), 'RPBK', $rowNumber);
                $totalValue = rowValue($row, $headers, 'TOTAL');
                $customers[$idpel] = [
                    'idpel' => $idpel,
                    'koded' => rowValue($row, $headers, 'KOKED'),
                    'nama' => $nama,
                    'alamat' => rowValue($row, $headers, 'ALAMAT'),
                    'tarif' => rowValue($row, $headers, 'TARIF'),
                    'daya' => parseAmount(rowValue($row, $headers, 'DAYA'), 'DAYA', $rowNumber),
                    'rptag' => $rptag,
                    'rpbk' => $rpbk,
                    'total' => $totalValue === '' ? $rptag + $rpbk : parseAmount($totalValue, 'TOTAL', $rowNumber),
                    'kontak_pelanggan' => rowValue($row, $headers, 'KONTAKPELANGGAN'),
                    'no_wa' => rowValue($row, $headers, 'NOWA')
                ];
            }
            if ($customers === []) {
                throw new RuntimeException('Tidak ada data pelanggan dengan IDPEL yang dapat diimpor.');
            }

            $conn->begin_transaction();
            try {
                $conn->query('DELETE FROM billing_customers');
                $stmt = $conn->prepare(
                    'INSERT INTO billing_customers
                    (idpel, koded, nama, alamat, tarif, daya, rptag, rpbk, total, kontak_pelanggan, no_wa, payment_status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, "unknown")'
                );
                foreach ($customers as $customer) {
                    $stmt->bind_param(
                        'sssssiiiiss',
                        $customer['idpel'],
                        $customer['koded'],
                        $customer['nama'],
                        $customer['alamat'],
                        $customer['tarif'],
                        $customer['daya'],
                        $customer['rptag'],
                        $customer['rpbk'],
                        $customer['total'],
                        $customer['kontak_pelanggan'],
                        $customer['no_wa']
                    );
                    $stmt->execute();
                }
                $stmt->close();
                setMeta($conn, 'master_imported_at', date('Y-m-d H:i:s'));
                setMeta($conn, 'daily_imported_at', '');
                $conn->commit();
            } catch (Throwable $exception) {
                $conn->rollback();
                error_log('Billing master import failed: ' . $exception->getMessage());
                throw new RuntimeException('Impor master gagal dan data sebelumnya tetap aman. Periksa log server untuk detailnya.');
            }

            setFlash('success', number_format(count($customers)) . ' data pelanggan berhasil dimuat. Status pembayaran menunggu file harian berikutnya.');
            redirectToPage();
        }

        if ($action === 'import_daily') {
            $rows = readTabularRows((string) $file['tmp_name'], $extension);
            if ($rows === []) {
                throw new RuntimeException('File IDPEL kosong atau tidak dapat dibaca.');
            }
            $headers = makeHeaderMap($rows[0]);
            $hasHeader = array_key_exists('IDPEL', $headers);
            $idpelIndex = $hasHeader ? $headers['IDPEL'] : 0;
            $idpels = [];
            $dailyRows = $hasHeader ? array_slice($rows, 1) : $rows;
            foreach ($dailyRows as $row) {
                $idpel = preg_replace('/\s+/', '', trim((string) ($row[$idpelIndex] ?? ''))) ?? '';
                if ($idpel !== '' && strtoupper($idpel) !== 'IDPEL') {
                    $idpels[$idpel] = true;
                }
            }
            if ($idpels === []) {
                throw new RuntimeException('Tidak ada IDPEL pada file harian. Pastikan kolom IDPEL berada di baris pertama.');
            }

            $conn->begin_transaction();
            try {
                $conn->query("UPDATE billing_customers SET payment_status = 'paid'");
                $stmt = $conn->prepare("UPDATE billing_customers SET payment_status = 'unpaid' WHERE idpel = ?");
                foreach (array_keys($idpels) as $idpel) {
                    $stmt->bind_param('s', $idpel);
                    $stmt->execute();
                }
                $stmt->close();
                setMeta($conn, 'daily_imported_at', date('Y-m-d H:i:s'));
                $conn->commit();
            } catch (Throwable $exception) {
                $conn->rollback();
                error_log('Billing daily import failed: ' . $exception->getMessage());
                throw new RuntimeException('Pembaruan status gagal dan status sebelumnya tetap aman. Periksa log server untuk detailnya.');
            }

            $stmt = $conn->prepare("SELECT COUNT(*) FROM billing_customers WHERE payment_status = 'unpaid'");
            $stmt->execute();
            $stmt->bind_result($matchedCount);
            $stmt->fetch();
            $stmt->close();
            setFlash(
                'success',
                'File harian berhasil diproses. ' . number_format((int) $matchedCount) .
                ' pelanggan dari master masih belum lunas; IDPEL yang tidak ada pada file ditandai lunas.'
            );
            redirectToPage();
        }

        throw new RuntimeException('Aksi tidak dikenali. Silakan muat ulang halaman.');
    } catch (RuntimeException $exception) {
        setFlash('error', $exception->getMessage());
        redirectToPage();
    }
}

$flash = $_SESSION['billing_flash'] ?? null;
unset($_SESSION['billing_flash']);
$search = trim((string) ($_GET['q'] ?? ''));
$statusFilter = (string) ($_GET['status'] ?? 'all');
if (!in_array($statusFilter, ['all', 'unpaid', 'paid', 'unknown'], true)) {
    $statusFilter = 'all';
}
$sortOptions = [
    'koded_asc' => ['label' => 'KODED A-Z', 'order' => 'bc.koded ASC, bc.nama ASC'],
    'koded_desc' => ['label' => 'KODED Z-A', 'order' => 'bc.koded DESC, bc.nama ASC'],
    'nama_asc' => ['label' => 'Nama A-Z', 'order' => 'bc.nama ASC'],
    'nama_desc' => ['label' => 'Nama Z-A', 'order' => 'bc.nama DESC'],
    'idpel_asc' => ['label' => 'IDPEL kecil-besar', 'order' => 'bc.idpel ASC'],
    'idpel_desc' => ['label' => 'IDPEL besar-kecil', 'order' => 'bc.idpel DESC'],
    'total_asc' => ['label' => 'Tagihan rendah-tinggi', 'order' => 'bc.total ASC, bc.nama ASC'],
    'total_desc' => ['label' => 'Tagihan tinggi-rendah', 'order' => 'bc.total DESC, bc.nama ASC']
];
$sort = (string) ($_GET['sort'] ?? 'koded_asc');
if (!isset($sortOptions[$sort])) {
    $sort = 'koded_asc';
}
$page = max(1, (int) ($_GET['page'] ?? 1));

$where = [];
$params = [];
$types = '';
if ($statusFilter !== 'all') {
    $where[] = 'bc.payment_status = ?';
    $params[] = $statusFilter;
    $types .= 's';
}
if ($search !== '') {
    $where[] = '(bc.idpel LIKE ? OR bc.koded LIKE ? OR bc.nama LIKE ? OR bc.alamat LIKE ? OR bc.no_wa LIKE ?)';
    $term = '%' . $search . '%';
    array_push($params, $term, $term, $term, $term, $term);
    $types .= 'sssss';
}
$whereSql = $where === [] ? '' : ' WHERE ' . implode(' AND ', $where);
$countStmt = $conn->prepare('SELECT COUNT(*) FROM billing_customers bc' . $whereSql);
if ($params !== []) {
    bindStatementParams($countStmt, $types, $params);
}
$countStmt->execute();
$countStmt->bind_result($filteredCount);
$countStmt->fetch();
$countStmt->close();
$filteredCount = (int) $filteredCount;
$totalPages = max(1, (int) ceil($filteredCount / PAGE_SIZE));
$page = min($page, $totalPages);
$offset = ($page - 1) * PAGE_SIZE;

$query = 'SELECT bc.idpel, bc.koded, bc.nama, bc.alamat, bc.tarif, bc.daya, bc.total,
                 bc.kontak_pelanggan, bc.no_wa, bc.payment_status, wc.clicked_on AS whatsapp_clicked_on
          FROM billing_customers bc
          LEFT JOIN billing_whatsapp_clicks wc ON wc.idpel = bc.idpel' . $whereSql .
          ' ORDER BY LEFT(bc.koded, 6) ASC, ' . $sortOptions[$sort]['order'] . ' LIMIT ? OFFSET ?';
$stmt = $conn->prepare($query);
$queryParams = $params;
$queryTypes = $types . 'ii';
$queryParams[] = PAGE_SIZE;
$queryParams[] = $offset;
bindStatementParams($stmt, $queryTypes, $queryParams);
$stmt->execute();
$customers = $stmt->get_result()->fetch_all(MYSQLI_ASSOC);
$stmt->close();

$stats = ['total' => 0, 'unpaid' => 0, 'paid' => 0, 'unknown' => 0];
$statsResult = $conn->query('SELECT payment_status, COUNT(*) AS count_rows FROM billing_customers GROUP BY payment_status');
while ($row = $statsResult->fetch_assoc()) {
    $stats[$row['payment_status']] = (int) $row['count_rows'];
    $stats['total'] += (int) $row['count_rows'];
}
$masterImportedAt = getMeta($conn, 'master_imported_at');
$dailyImportedAt = getMeta($conn, 'daily_imported_at');

function formatImportedDate(string $value): string
{
    if ($value === '') {
        return 'Belum pernah';
    }
    $timestamp = strtotime($value);
    return $timestamp === false ? $value : date('d M Y, H:i', $timestamp);
}

function pageUrl(int $page, string $search, string $status, string $sort): string
{
    return '?' . http_build_query(['q' => $search, 'status' => $status, 'sort' => $sort, 'page' => $page]);
}
?>
<!doctype html>
<html lang="id">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="theme-color" content="#102a43">
    <title>Tagihan Pelanggan | KTS</title>
    <script>
        try {
            document.documentElement.dataset.theme = localStorage.getItem('kts-theme') === 'dark' ? 'dark' : 'light';
        } catch (error) {
            document.documentElement.dataset.theme = 'light';
        }
    </script>
    <link rel="stylesheet" href="style.css?v=<?= (int) filemtime(__DIR__ . '/style.css') ?>">
</head>
<body>
<header class="topbar">
    <a class="brand" href="index.php" aria-label="KTS Beranda">
        <span class="brand-mark">K</span>
        <span><strong>KTS</strong><small>LAYANAN TAGIHAN PELANGGAN</small></span>
    </a>
    <div class="topbar-actions">
        <div class="topbar-note"><span class="live-dot"></span> Panel pengelolaan tagihan</div>
        <button class="theme-toggle" id="theme-toggle" type="button" aria-label="Aktifkan tema gelap" aria-pressed="false">
            <span class="theme-icon" aria-hidden="true">☾</span><span class="theme-label">Gelap</span>
        </button>
    </div>
</header>

<main class="page-shell">
    <section class="welcome">
        <div>
            <p class="eyebrow">DASHBOARD OPERASIONAL</p>
            <h1>Tagihan pelanggan</h1>
            <p class="welcome-copy">Kelola data master, pantau status pembayaran, dan kirim pengingat melalui WhatsApp.</p>
        </div>
        <div class="welcome-date">
            <span class="date-icon" aria-hidden="true">▦</span>
            <span><small>HARI INI</small><strong><?= h(date('d M Y')) ?></strong></span>
        </div>
    </section>

    <?php if (is_array($flash)): ?>
        <div class="alert alert-<?= h((string) $flash['type']) ?>" role="status">
            <span class="alert-icon"><?= $flash['type'] === 'success' ? '✓' : '!' ?></span>
            <span><?= h((string) $flash['message']) ?></span>
        </div>
    <?php endif; ?>

    <section class="stats-grid" aria-label="Ringkasan pelanggan">
        <article class="stat-card stat-total">
            <span class="stat-icon">◎</span>
            <div><p>Total pelanggan</p><strong><?= number_format($stats['total']) ?></strong><small>Data dalam master</small></div>
        </article>
        <article class="stat-card stat-unpaid">
            <span class="stat-icon">◷</span>
            <div><p>Belum lunas</p><strong><?= number_format($stats['unpaid']) ?></strong><small>Masih ada di data harian</small></div>
        </article>
        <article class="stat-card stat-paid">
            <span class="stat-icon">✓</span>
            <div><p>Sudah lunas</p><strong><?= number_format($stats['paid']) ?></strong><small>Tidak ada di data harian</small></div>
        </article>
        <article class="stat-card stat-unknown">
            <span class="stat-icon">—</span>
            <div><p>Belum diperiksa</p><strong><?= number_format($stats['unknown']) ?></strong><small>Menunggu file harian</small></div>
        </article>
    </section>

    <section class="import-grid" aria-label="Impor data">
        <article class="panel import-panel">
            <div class="panel-heading">
                <span class="panel-icon master-icon">↥</span>
                <div><h2>Data master bulanan</h2><p>Ganti seluruh data pelanggan dari file Excel terbaru.</p></div>
            </div>
            <form method="post" enctype="multipart/form-data" class="upload-form" data-confirm="Data master saat ini akan diganti seluruhnya. Lanjutkan?">
                <input type="hidden" name="csrf_token" value="<?= h($_SESSION['billing_csrf']) ?>">
                <input type="hidden" name="action" value="import_master">
                <label class="file-picker">
                    <span class="file-picker-icon">＋</span>
                    <span class="file-picker-text"><strong>Pilih file master</strong><small>Format .xlsx · Maks. 10 MB</small></span>
                    <input type="file" name="spreadsheet" accept=".xlsx" required>
                </label>
                <button class="button button-primary" type="submit">Ganti data master <span>→</span></button>
            </form>
            <div class="panel-foot"><span>Kolom IDPEL dan NAMA wajib tersedia.</span><span>Terakhir: <?= h(formatImportedDate($masterImportedAt)) ?></span></div>
        </article>

        <article class="panel import-panel">
            <div class="panel-heading">
                <span class="panel-icon daily-icon">↻</span>
                <div><h2>Pembaruan status harian</h2><p>Unggah daftar IDPEL yang masih memiliki tagihan.</p></div>
            </div>
            <form method="post" enctype="multipart/form-data" class="upload-form">
                <input type="hidden" name="csrf_token" value="<?= h($_SESSION['billing_csrf']) ?>">
                <input type="hidden" name="action" value="import_daily">
                <label class="file-picker">
                    <span class="file-picker-icon">＋</span>
                    <span class="file-picker-text"><strong>Pilih file IDPEL</strong><small>Format .xlsx atau .csv · Maks. 10 MB</small></span>
                    <input type="file" name="spreadsheet" accept=".xlsx,.csv" required>
                </label>
                <button class="button button-secondary" type="submit">Perbarui status <span>→</span></button>
            </form>
            <div class="panel-foot"><span>IDPEL yang tidak tercantum akan ditandai lunas.</span><span>Terakhir: <?= h(formatImportedDate($dailyImportedAt)) ?></span></div>
        </article>
    </section>

    <section class="panel customer-panel">
        <div class="customer-heading">
            <div><p class="eyebrow">DATA PELANGGAN</p><h2>Daftar tagihan</h2><p class="section-description">Cari pelanggan, tinjau status, lalu buka WhatsApp untuk menyiapkan pesan tagihan.</p></div>
            <div class="table-count"><strong><?= number_format($filteredCount) ?></strong> pelanggan</div>
        </div>

        <form class="toolbar" method="get" action="index.php">
            <label class="search-box">
                <span aria-hidden="true">⌕</span>
                <input type="search" name="q" value="<?= h($search) ?>" placeholder="Cari nama, IDPEL, KODED, alamat..." aria-label="Cari pelanggan">
            </label>
            <label class="filter-label" for="status-filter">Status</label>
            <select id="status-filter" name="status">
                <option value="all" <?= $statusFilter === 'all' ? 'selected' : '' ?>>Semua status</option>
                <option value="unpaid" <?= $statusFilter === 'unpaid' ? 'selected' : '' ?>>Belum lunas</option>
                <option value="paid" <?= $statusFilter === 'paid' ? 'selected' : '' ?>>Lunas</option>
                <option value="unknown" <?= $statusFilter === 'unknown' ? 'selected' : '' ?>>Belum diperiksa</option>
            </select>
            <label class="filter-label" for="sort-filter">Urutkan</label>
            <select id="sort-filter" name="sort">
                <?php foreach ($sortOptions as $value => $option): ?>
                    <option value="<?= h($value) ?>" <?= $sort === $value ? 'selected' : '' ?>><?= h($option['label']) ?></option>
                <?php endforeach; ?>
            </select>
            <button class="button button-search" type="submit">Terapkan</button>
        </form>

        <div class="table-wrap">
            <table>
                <thead><tr>
                    <th>Pelanggan</th><th>IDPEL / KODED</th><th>Tarif / daya</th><th>Tagihan</th><th>Status pembayaran</th><th>Aksi</th>
                </tr></thead>
                <tbody>
                <?php if ($customers === []): ?>
                    <tr><td class="empty-state" colspan="6"><span>⌕</span><strong>Belum ada data pelanggan</strong><small>Unggah data master atau ubah kata kunci pencarian Anda.</small></td></tr>
                <?php else: ?>
                    <?php $currentGroup = null; ?>
                    <?php foreach ($customers as $customer):
                        $groupKey = substr(trim((string) $customer['koded']), 0, 6);
                        if ($groupKey !== $currentGroup):
                            $currentGroup = $groupKey;
                    ?>
                        <tr class="group-row"><td colspan="6"><span class="group-label">GRUP KODED</span><strong><?= $groupKey !== '' ? h($groupKey) : 'Kode belum tersedia' ?></strong><small><?= $groupKey !== '' ? '6 karakter pertama' : 'Lengkapi KOKED pada data master untuk mengelompokkan pelanggan ini.' ?></small></td></tr>
                    <?php endif;
                        $phone = normalizeWhatsapp((string) ($customer['no_wa'] !== '' ? $customer['no_wa'] : $customer['kontak_pelanggan']));
                        $whatsappClickedToday = ($customer['whatsapp_clicked_on'] ?? '') === date('Y-m-d');
                        $statusLabels = ['unpaid' => 'Belum lunas', 'paid' => 'Lunas', 'unknown' => 'Belum diperiksa'];
                        $message = "Yth. Bpk/Ibu {$customer['nama']},\n\nBerikut informasi tagihan PLN Anda:\nIDPEL: {$customer['idpel']}\nTotal tagihan: " . formatRupiah($customer['total']) . "\n\nMohon melakukan pembayaran sebelum tanggal 20 setiap bulannya. Terima kasih.";
                        $whatsappUrl = $phone !== '' ? 'https://wa.me/' . rawurlencode($phone) . '?text=' . rawurlencode($message) : '';
                    ?>
                        <tr>
                            <td class="customer-cell">
                                <span class="avatar"><?= h(strtoupper(substr((string) $customer['nama'], 0, 1))) ?></span>
                                <span><strong><?= h((string) $customer['nama']) ?></strong><small><?= h((string) $customer['alamat']) ?></small></span>
                            </td>
                            <td class="id-cell">
                                <span class="id-number"><?= h((string) $customer['idpel']) ?></span>
                                <small class="koded-cell"><span>KODED</span> <?= h((string) ($customer['koded'] !== '' ? $customer['koded'] : '—')) ?></small>
                            </td>
                            <td><span class="rate"><?= h((string) $customer['tarif']) ?></span><small class="sub-cell"><?= number_format((int) $customer['daya']) ?> VA</small></td>
                            <td class="amount-cell"><?= h(formatRupiah($customer['total'])) ?></td>
                            <td><span class="status-badge status-<?= h((string) $customer['payment_status']) ?>"><i></i><?= h($statusLabels[$customer['payment_status']] ?? 'Belum diperiksa') ?></span></td>
                            <td>
                                <?php if ($customer['payment_status'] === 'unpaid' && $whatsappUrl !== ''): ?>
                                    <a class="whatsapp-button<?= $whatsappClickedToday ? ' whatsapp-button-clicked' : '' ?>" href="<?= h($whatsappUrl) ?>" target="_blank" rel="noopener noreferrer" aria-label="Siapkan pesan WhatsApp untuk <?= h((string) $customer['nama']) ?>" data-idpel="<?= h((string) $customer['idpel']) ?>" data-csrf-token="<?= h($_SESSION['billing_csrf']) ?>" data-clicked-on="<?= h((string) ($customer['whatsapp_clicked_on'] ?? '')) ?>" title="<?= $whatsappClickedToday ? 'WhatsApp sudah dibuka hari ini; klik untuk membuka kembali.' : 'Kirim invoice melalui WhatsApp.' ?>">
                                        <span class="whatsapp-mark">WA</span><span>Kirim via WhatsApp</span>
                                    </a>
                                <?php elseif ($customer['payment_status'] === 'paid'): ?>
                                    <span class="no-phone">Tidak perlu ditagih</span>
                                <?php elseif ($customer['payment_status'] === 'unknown'): ?>
                                    <span class="no-phone">Periksa status dulu</span>
                                <?php elseif ($whatsappUrl === ''): ?>
                                    <span class="no-phone">Nomor tidak tersedia</span>
                                <?php else: ?>
                                    <span class="no-phone">—</span>
                                <?php endif; ?>
                            </td>
                        </tr>
                    <?php endforeach; ?>
                <?php endif; ?>
                </tbody>
            </table>
        </div>

        <?php if ($totalPages > 1): ?>
            <nav class="pagination" aria-label="Navigasi halaman">
                <span>Halaman <?= number_format($page) ?> dari <?= number_format($totalPages) ?></span>
                <div>
                    <?php if ($page > 1): ?><a href="<?= h(pageUrl($page - 1, $search, $statusFilter, $sort)) ?>">← Sebelumnya</a><?php endif; ?>
                    <?php if ($page < $totalPages): ?><a href="<?= h(pageUrl($page + 1, $search, $statusFilter, $sort)) ?>">Berikutnya →</a><?php endif; ?>
                </div>
            </nav>
        <?php endif; ?>
    </section>

    <p class="page-note"><span>i</span> Pesan WhatsApp akan terbuka dengan isi invoice yang sudah disiapkan. Periksa kembali, lalu tekan <strong>Kirim</strong> di WhatsApp. Warna tombol berubah saat WhatsApp dibuka; ini bukan konfirmasi pesan terkirim.</p>
</main>
<footer class="footer"><span>KTS · Panel Tagihan Pelanggan</span><span>Data pembayaran berdasarkan file harian terakhir</span></footer>
<script>
const themeToggle = document.getElementById('theme-toggle');
const themeColor = document.querySelector('meta[name="theme-color"]');

function updateThemeButton() {
    const isDark = document.documentElement.dataset.theme === 'dark';
    themeToggle.setAttribute('aria-pressed', String(isDark));
    themeToggle.setAttribute('aria-label', isDark ? 'Aktifkan tema terang' : 'Aktifkan tema gelap');
    themeToggle.querySelector('.theme-icon').textContent = isDark ? '☼' : '☾';
    themeToggle.querySelector('.theme-label').textContent = isDark ? 'Terang' : 'Gelap';
    themeColor.setAttribute('content', isDark ? '#101b25' : '#102a43');
}

updateThemeButton();
themeToggle.addEventListener('click', () => {
    const nextTheme = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = nextTheme;
    updateThemeButton();
    try {
        localStorage.setItem('kts-theme', nextTheme);
    } catch (error) {
        console.warn('Preferensi tema tidak dapat disimpan di browser.', error);
    }
});

document.querySelectorAll('.whatsapp-button[data-idpel]').forEach((button) => {
    button.addEventListener('click', async () => {
        const wasAlreadyClickedToday = button.classList.contains('whatsapp-button-clicked');
        button.classList.add('whatsapp-button-clicked');
        button.dataset.clickedOn = getJakartaDate();
        button.title = 'WhatsApp sudah dibuka hari ini; klik untuk membuka kembali.';

        try {
            const response = await fetch(window.location.pathname, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
                body: new URLSearchParams({
                    action: 'mark_whatsapp',
                    idpel: button.dataset.idpel,
                    csrf_token: button.dataset.csrfToken
                })
            });
            const result = await response.json();
            if (!response.ok || !result.success) {
                throw new Error(result.message || 'Penanda WhatsApp tidak dapat disimpan.');
            }
        } catch (error) {
            if (!wasAlreadyClickedToday) {
                button.classList.remove('whatsapp-button-clicked');
                button.title = 'Kirim invoice melalui WhatsApp.';
            }
            console.error('Gagal menyimpan penanda WhatsApp.', error);
            window.alert('WhatsApp terbuka, tetapi penanda warna tidak dapat disimpan. Coba muat ulang halaman atau hubungi administrator.');
        }
    });
});

function getJakartaDate() {
    const parts = new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Asia/Jakarta',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
    }).formatToParts(new Date());
    const dateParts = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    return `${dateParts.year}-${dateParts.month}-${dateParts.day}`;
}

setInterval(() => {
    const today = getJakartaDate();
    document.querySelectorAll('.whatsapp-button-clicked').forEach((button) => {
        if (button.dataset.clickedOn !== today) {
            button.classList.remove('whatsapp-button-clicked');
            button.title = 'Kirim invoice melalui WhatsApp.';
        }
    });
}, 60000);

document.querySelectorAll('.file-picker input[type="file"]').forEach((input) => {
    input.addEventListener('change', () => {
        const label = input.closest('.file-picker');
        const name = input.files && input.files[0] ? input.files[0].name : '';
        const title = label.querySelector('.file-picker-text strong');
        if (name) title.textContent = name;
    });
});
document.querySelectorAll('form[data-confirm]').forEach((form) => {
    form.addEventListener('submit', (event) => {
        if (!window.confirm(form.dataset.confirm)) event.preventDefault();
    });
});
</script>
</body>
</html>
