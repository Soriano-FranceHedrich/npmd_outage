<?php
// Outage Log request handlers. Included by index.php during POST requests.

    if ($postAction === 'upload_outages') {
        $upload = $_FILES['outage_csv'] ?? null;
        if (!$upload || !is_uploaded_file($upload['tmp_name'] ?? '')) {
            flash('error', 'Please choose an outage file to import.');
            redirect('?action=outages');
        }

        $rows = parse_uploaded_file($upload['tmp_name'], $upload['name'] ?? '');
        if (!$rows) {
            flash('error', 'The uploaded file is empty or unreadable. Please upload a CSV, XLS, or XLSX file.');
            redirect('?action=outages');
        }
        $firstRowKeys = array_keys($rows[0]);
        $hasOutageHeaders = in_array('down_time', $firstRowKeys, true)
            && (bool)array_intersect(['job_order_no', 'job_order', 'project_name', 'project_office', 'project', 'office', 'office_name'], $firstRowKeys);
        if (!$hasOutageHeaders) {
            // Some spreadsheet editors alter header cells when converting the system's XLS export to XLSX.
            // The system export's column order is stable, so use it as a fallback for that workbook format.
            $exportColumns = ['project_office', 'region', 'team', 'status', 'priority', 'root_cause', 'down_time', 'up_time', 'duration', 'remarks'];
            $rows = array_map(static function ($row) use ($exportColumns) {
                $values = array_values($row);
                $mapped = [];
                foreach ($exportColumns as $index => $key) {
                    $mapped[$key] = trim((string)($values[$index] ?? ''));
                }
                return $mapped;
            }, $rows);
        }
        $_SESSION['pending_outage_imports'] = $rows;
        redirect('?action=outages&import_review=1');
    }

    if ($postAction === 'cancel_import_outages') {
        unset($_SESSION['pending_outage_imports']);
        redirect('?action=outages');
    }

    if ($postAction === 'confirm_import_outages') {
        $submitted = $_POST['outages'] ?? [];
        $submitted = is_array($submitted) ? $submitted : [];
        $_SESSION['pending_outage_imports'] = array_map(static function ($row) {
            $row = is_array($row) ? $row : [];
            return [
                'project_id' => trim((string)($row['project_id'] ?? '')),
                'status' => trim((string)($row['status'] ?? 'Open')),
                'priority' => trim((string)($row['priority'] ?? 'Normal')),
                'root_cause' => trim((string)($row['root_cause'] ?? 'Unknown')),
                'down_time' => trim((string)($row['down_time'] ?? '')),
                'up_time' => trim((string)($row['up_time'] ?? '')),
                'remarks' => trim((string)($row['remarks'] ?? '')),
            ];
        }, $submitted);

        $outagesToSave = [];
        $allowedStatuses = ['Open', 'Ongoing', 'Restored', 'Closed'];
        $allowedPriorities = ['Low', 'Normal', 'High', 'Critical'];
        $allowedCauses = ['Power Failure', 'Equipment Failure', 'Maintenance Activity', 'Weather', 'Configuration Error', 'Third-Party Damage', 'Unknown', 'Other'];
        foreach ($submitted as $index => $row) {
            $row = is_array($row) ? $row : [];
            $projectId = (int)($row['project_id'] ?? 0);
            $downTime = normalize_import_datetime((string)($row['down_time'] ?? ''));
            $upTime = normalize_import_datetime((string)($row['up_time'] ?? ''));
            $status = trim((string)($row['status'] ?? 'Open'));
            $priority = trim((string)($row['priority'] ?? 'Normal'));
            $rootCause = trim((string)($row['root_cause'] ?? 'Unknown'));
            if ($projectId <= 0 || $downTime === null) {
                flash('error', 'Every imported outage needs a project and Down Time. No records were saved.');
                redirect('?action=outages&import_review=1');
            }
            if ($upTime !== null && strtotime($upTime) < strtotime($downTime)) {
                flash('error', 'Up Time cannot be earlier than Down Time. No records were saved.');
                redirect('?action=outages&import_review=1');
            }
            if (!in_array($status, $allowedStatuses, true) || !in_array($priority, $allowedPriorities, true) || !in_array($rootCause, $allowedCauses, true)) {
                flash('error', 'An imported outage has an unsupported Status, Priority, or Root Cause. Review the fields and try again.');
                redirect('?action=outages&import_review=1');
            }
            $upTime = normalize_outage_up_time($downTime, $upTime, $status);
            $outagesToSave[] = [
                $projectId,
                $downTime,
                $upTime,
                $status,
                $priority,
                $rootCause,
                trim((string)($row['remarks'] ?? '')),
            ];
        }
        if (!$outagesToSave) {
            flash('error', 'There are no outage records to save.');
            redirect('?action=outages');
        }

        $pdo->beginTransaction();
        try {
            $projectQuery = $pdo->prepare('SELECT province, project_name FROM projects WHERE id=?');
            $insert = $pdo->prepare('INSERT INTO outages (project_id, down_time, up_time, status, priority, root_cause, remarks) VALUES (?, ?, ?, ?, ?, ?, ?)');
            $setCode = $pdo->prepare('UPDATE outages SET outage_code=? WHERE id=?');
            foreach ($outagesToSave as $outageData) {
                $projectQuery->execute([$outageData[0]]);
                $projectRow = $projectQuery->fetch();
                if (!$projectRow) {
                    throw new RuntimeException('An imported outage references a project that no longer exists.');
                }
                $insert->execute($outageData);
                $outageId = (int)$pdo->lastInsertId();
                $setCode->execute([outage_code((string)$projectRow['province'], $outageId, (string)$projectRow['project_name']), $outageId]);
            }
            $pdo->commit();
        } catch (Throwable $e) {
            if ($pdo->inTransaction()) $pdo->rollBack();
            if ($e instanceof PDOException && (stripos($e->getMessage(), 'Duplicate entry') !== false || stripos($e->getMessage(), 'uq_outages_outage_code') !== false)) {
                flash('error', 'An outage code already exists. No records were saved; please review the import and try again.');
                redirect('?action=outages&import_review=1');
            }
            throw $e;
        }
        unset($_SESSION['pending_outage_imports']);
        flash('success', 'All ' . count($outagesToSave) . ' imported outage record(s) saved.');
        redirect('?action=outages');
    }

    if ($postAction === 'save_outage') {
        $projectId = (int)($_POST['project_id'] ?? 0);
        $down = normalize_datetime_input($_POST['down_time'] ?? '');
        $up = normalize_datetime_input($_POST['up_time'] ?? '');
        $status = trim($_POST['status'] ?? 'Open');
        $up = normalize_outage_up_time($down, $up, $status);
        $priority = trim($_POST['priority'] ?? 'Normal');
        $rootCause = trim($_POST['root_cause'] ?? 'Unknown');
        $remarks = trim($_POST['remarks'] ?? '');

        if ($projectId <= 0 || $down === null || $down === '') {
            flash('error', 'Project and Down Time are required.');
            redirect('?action=outages');
        }

        $upValue = $up !== null && $up !== '' ? $up : null;
        $stmt = $pdo->prepare("INSERT INTO outages (project_id, down_time, up_time, status, priority, root_cause, remarks)
                               VALUES (?, ?, ?, ?, ?, ?, ?)");
        $stmt->execute([$projectId, $down, $upValue, $status, $priority, $rootCause, $remarks]);
        $outageId = (int)$pdo->lastInsertId();
        $projectRow = $pdo->query("SELECT province, project_name FROM projects WHERE id=" . $projectId)->fetch(PDO::FETCH_ASSOC);
        $pdo->prepare("UPDATE outages SET outage_code=? WHERE id=?")
            ->execute([
                outage_code((string)($projectRow['province'] ?? ''), $outageId, (string)($projectRow['project_name'] ?? '')),
                $outageId
            ]);

        flash('success', $upValue ? 'Outage recorded.' : 'Outage opened. Add the Up Time when service returns.');
        redirect('?action=outages');
    }

    if ($postAction === 'update_outage') {
        $id = (int)($_POST['id'] ?? 0);
        $projectId = (int)($_POST['project_id'] ?? 0);
        $down = normalize_datetime_input($_POST['down_time'] ?? '');
        $up = normalize_datetime_input($_POST['up_time'] ?? '');
        $status = trim($_POST['status'] ?? 'Open');
        $up = normalize_outage_up_time($down, $up, $status);
        $priority = trim($_POST['priority'] ?? 'Normal');
        $rootCause = trim($_POST['root_cause'] ?? 'Unknown');
        $remarks = trim($_POST['remarks'] ?? '');

        if ($id <= 0 || $projectId <= 0) {
            flash('error', 'Project is required.');
            redirect('?action=outages');
        }

        $upValue = $up !== null && $up !== '' ? $up : null;
        $stmt = $pdo->prepare("UPDATE outages
                               SET project_id=?, down_time=?, up_time=?, status=?, priority=?, root_cause=?, remarks=?
                               WHERE id=?");
        $stmt->execute([$projectId, $down, $upValue, $status, $priority, $rootCause, $remarks, $id]);

        $projectRow = $pdo->query("SELECT province, project_name FROM projects WHERE id=" . $projectId)->fetch(PDO::FETCH_ASSOC);
        $pdo->prepare("UPDATE outages SET outage_code=? WHERE id=?")
            ->execute([
                outage_code((string)($projectRow['province'] ?? ''), $id, (string)($projectRow['project_name'] ?? '')),
                $id
            ]);

        flash('success', 'Outage record updated.');
        redirect('?action=outages');
    }

    if ($postAction === 'close_outage') {
        $id = (int)($_POST['id'] ?? 0);
        $up = normalize_datetime_input($_POST['up_time'] ?? date('Y-m-d H:i'));
        if ($id > 0 && $up !== null && $up !== '') {
            $pdo->prepare("UPDATE outages SET up_time=?, status='Closed' WHERE id=? AND up_time IS NULL")
                ->execute([$up, $id]);
            flash('success', 'Outage closed.');
        }
        redirect('?action=outages');
    }

    if ($postAction === 'delete_outage') {
        $id = (int)($_POST['id'] ?? 0);
        $pdo->prepare("DELETE FROM outages WHERE id=?")->execute([$id]);
        flash('success', 'Outage deleted.');
        redirect('?action=outages');
    }
