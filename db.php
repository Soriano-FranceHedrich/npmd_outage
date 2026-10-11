<?php
declare(strict_types=1);
require_once __DIR__ . '/config.php';

function db(): PDO {
    static $pdo = null;
    if ($pdo instanceof PDO) return $pdo;

    $dsn = sprintf('mysql:host=%s;port=%d;dbname=%s;charset=%s', DB_HOST, DB_PORT, DB_NAME, DB_CHARSET);
    $pdo = new PDO($dsn, DB_USER, DB_PASS, [
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
    ]);

    $pdo->exec("CREATE TABLE IF NOT EXISTS projects (
        id INT NOT NULL AUTO_INCREMENT,
        job_order_no VARCHAR(255) NOT NULL DEFAULT '',
        project_name VARCHAR(255) NOT NULL,
        province VARCHAR(255) NOT NULL DEFAULT '',
        region VARCHAR(255) NOT NULL DEFAULT '',
        location VARCHAR(255) NOT NULL DEFAULT '',
        assigned_team VARCHAR(255) NOT NULL DEFAULT '',
        remarks TEXT NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_projects_job_order_no (job_order_no)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4");

    $pdo->exec("CREATE TABLE IF NOT EXISTS outages (
        id INT NOT NULL AUTO_INCREMENT,
        outage_code VARCHAR(32) NULL,
        project_id INT NOT NULL,
        down_time DATETIME NOT NULL,
        up_time DATETIME NULL,
        status VARCHAR(50) NOT NULL DEFAULT 'Open',
        priority VARCHAR(50) NOT NULL DEFAULT 'Normal',
        root_cause VARCHAR(100) NOT NULL DEFAULT 'Unknown',
        remarks TEXT NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        KEY idx_outages_project (project_id),
        KEY idx_outages_down (down_time),
        CONSTRAINT fk_outages_project FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4");

    $outageColumns = $pdo->query("SHOW COLUMNS FROM outages")->fetchAll(PDO::FETCH_ASSOC);
    foreach ([
        ['outage_code', "ALTER TABLE outages ADD COLUMN outage_code VARCHAR(32) NULL AFTER id"],
        ['status', "ALTER TABLE outages ADD COLUMN status VARCHAR(50) NOT NULL DEFAULT 'Open'"],
        ['priority', "ALTER TABLE outages ADD COLUMN priority VARCHAR(50) NOT NULL DEFAULT 'Normal'"],
        ['root_cause', "ALTER TABLE outages ADD COLUMN root_cause VARCHAR(100) NOT NULL DEFAULT 'Unknown'"]
    ] as [$columnName, $sql]) {
        $exists = false;
        foreach ($outageColumns as $column) {
            if ($column['Field'] === $columnName) {
                $exists = true;
                break;
            }
        }
        if (!$exists) {
            $pdo->exec($sql);
        }
    }

    // Earlier date-picker code copied Down Time into an empty Up Time field.
    // For active records where the two timestamps are equal, restore the open state.
    $pdo->exec("UPDATE outages SET up_time = NULL
        WHERE LOWER(TRIM(status)) IN ('open', 'ongoing')
          AND up_time IS NOT NULL
          AND up_time = down_time");

    $projectColumns = $pdo->query("SHOW COLUMNS FROM projects")->fetchAll(PDO::FETCH_ASSOC);
    $hasProvince = false;
    foreach ($projectColumns as $column) {
        if ($column['Field'] === 'province') {
            $hasProvince = true;
            break;
        }
    }
    if (!$hasProvince) {
        $pdo->exec("ALTER TABLE projects ADD COLUMN province VARCHAR(255) NOT NULL DEFAULT '' AFTER project_name");
    }

    $outagesForCodeMigration = $pdo->query("SELECT o.id, p.province, p.project_name
        FROM outages o
        JOIN projects p ON p.id=o.project_id")->fetchAll(PDO::FETCH_ASSOC);
    $updateOutageCode = $pdo->prepare("UPDATE outages SET outage_code=? WHERE id=?");
    foreach ($outagesForCodeMigration as $outage) {
        $updateOutageCode->execute([
            outage_code((string)$outage['province'], (int)$outage['id'], (string)$outage['project_name']),
            (int)$outage['id']
        ]);
    }

    $outageCodeIndexExists = $pdo->query("SELECT COUNT(*)
        FROM information_schema.statistics
        WHERE table_schema = DATABASE()
          AND table_name = 'outages'
          AND index_name = 'uq_outages_outage_code'")->fetchColumn();
    if (!$outageCodeIndexExists) {
        $pdo->exec("ALTER TABLE outages ADD UNIQUE KEY uq_outages_outage_code (outage_code)");
    }

    $columns = $pdo->query("SHOW COLUMNS FROM projects")->fetchAll(PDO::FETCH_ASSOC);
    $hasJobOrderNo = false;
    foreach ($columns as $column) {
        if ($column['Field'] === 'job_order_no') {
            $hasJobOrderNo = true;
            break;
        }
    }
    if (!$hasJobOrderNo) {
        $pdo->exec("ALTER TABLE projects ADD COLUMN job_order_no VARCHAR(255) NOT NULL DEFAULT ''");
    }

    $duplicateJOs = $pdo->query("SELECT job_order_no, COUNT(*) AS total
        FROM projects
        WHERE TRIM(job_order_no) <> ''
        GROUP BY job_order_no
        HAVING COUNT(*) > 1")->fetchAll(PDO::FETCH_ASSOC);

    $indexExists = $pdo->query("SELECT COUNT(*)
        FROM information_schema.statistics
        WHERE table_schema = DATABASE()
          AND table_name = 'projects'
          AND index_name = 'uq_projects_job_order_no'")->fetchColumn();

    if (!$indexExists && !$duplicateJOs) {
        $pdo->exec("ALTER TABLE projects ADD UNIQUE KEY uq_projects_job_order_no (job_order_no)");
    }

    $pdo->exec("UPDATE projects SET job_order_no = CONCAT('JO-', LPAD(CAST(id AS CHAR), 6, '0')) WHERE TRIM(job_order_no) = ''");

    return $pdo;
}

function h(?string $value): string {
    return htmlspecialchars($value ?? '', ENT_QUOTES, 'UTF-8');
}

function redirect(string $url): never {
    header('Location: ' . $url);
    exit;
}

function flash(string $type, string $message): void {
    $_SESSION['flash'] = ['type' => $type, 'message' => $message];
}

function get_flash(): ?array {
    $f = $_SESSION['flash'] ?? null;
    unset($_SESSION['flash']);
    return $f;
}

function normalize_datetime_input(?string $value): ?string {
    $value = trim((string)($value ?? ''));
    if ($value === '') return null;
    $value = str_replace('T', ' ', $value);
    if (preg_match('/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/', $value)) {
        return $value . ':00';
    }
    return $value;
}

function normalize_outage_up_time(?string $downTime, ?string $upTime, string $status): ?string {
    if ($upTime !== null && $downTime !== null
        && in_array(strtolower(trim($status)), ['open', 'ongoing'], true)
        && strtotime($upTime) === strtotime($downTime)) {
        return null;
    }
    return $upTime;
}

function format_dt(?string $value): string {
    if (!$value) return '—';
    $t = strtotime($value);
    return $t ? date('M d, Y h:i A', $t) : $value;
}

function outage_code(string $province, int $outageId, ?string $projectName = null): string {
    $tokens = [$province, $projectName ?? ''];
    $safeCode = '';
    foreach ($tokens as $token) {
        $candidate = strtoupper(trim((string)preg_replace('/[^A-Za-z0-9]+/', '-', $token), '-'));
        if ($candidate !== '') {
            $safeCode = $candidate;
            break;
        }
    }
    $safeCode = $safeCode !== '' ? $safeCode : 'NPMD';
    return 'OUT-' . $safeCode . '-' . str_pad((string)$outageId, 6, '0', STR_PAD_LEFT);
}
