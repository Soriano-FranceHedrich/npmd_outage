<?php
// Project Offices request handlers. Included by index.php during POST requests.

    // Keep JO validation and writes serialized even on older databases where
    // the unique index could not be installed because legacy duplicates exist.
    if (in_array($postAction, ['upload_projects', 'confirm_import_project', 'save_project', 'delete_project'], true)) {
        $lock = $pdo->prepare("SELECT GET_LOCK('npmd_projects_job_order_write', 10)");
        $lock->execute();
        if ((int)$lock->fetchColumn() !== 1) {
            flash('error', 'Projects are being saved right now. Please try again.');
            redirect('?action=projects');
        }
    }

    if ($postAction === 'upload_projects') {
        $upload = $_FILES['project_csv'] ?? null;
        if (!$upload || !is_uploaded_file($upload['tmp_name'] ?? '')) {
            flash('error', 'Please choose a CSV file to import.');
            redirect('?action=projects');
        }

        $rows = parse_uploaded_file($upload['tmp_name'], $upload['name'] ?? '');
        if (!$rows) {
            flash('error', 'The uploaded file is empty or not readable. Please upload a CSV or .xlsx file.');
            redirect('?action=projects');
        }

        $extension = strtolower(pathinfo($upload['name'] ?? '', PATHINFO_EXTENSION));
        if (in_array($extension, ['xls', 'xlsx'], true)) {
            $_SESSION['pending_project_imports'] = $rows;
            redirect('?action=projects&import_review=1');
        }

        $imported = 0;
        $skipped = 0;
        $issues = [];
        $existingJobOrders = [];
        $existingProjectNames = [];
        foreach ($pdo->query('SELECT job_order_no, project_name FROM projects') as $existingProject) {
            $existingJobOrders[normalize_job_order_no((string)$existingProject['job_order_no'])] = true;
            $existingProjectNames[normalize_project_name((string)$existingProject['project_name'])] = true;
        }
        foreach ($rows as $rowIndex => $row) {
            $rowNumber = $rowIndex + 2;
            $jobOrderNo = trim((string)($row['job_order_no'] ?? $row['job_order'] ?? $row['job_order_no_'] ?? ''));
            $projectName = trim((string)($row['project_name'] ?? $row['project'] ?? $row['project_office'] ?? $row['office_name'] ?? ''));
            if ($jobOrderNo === '' || $projectName === '') {
                $skipped++;
                $missing = [];
                if ($jobOrderNo === '') $missing[] = 'Job Order No.';
                if ($projectName === '') $missing[] = 'Project Name';
                $issues[] = 'Row ' . $rowNumber . ': missing ' . implode(' and ', $missing) . '.';
                continue;
            }

            $normalizedJo = normalize_job_order_no($jobOrderNo);
            $normalizedName = normalize_project_name($projectName);
            if (isset($existingJobOrders[$normalizedJo])) {
                $skipped++;
                $issues[] = 'Row ' . $rowNumber . ': Job Order No. "' . $jobOrderNo . '" already exists (case-insensitive match).';
                continue;
            }
            if (isset($existingProjectNames[$normalizedName])) {
                $skipped++;
                $issues[] = 'Row ' . $rowNumber . ': Project Name "' . $projectName . '" already exists (case-insensitive match).';
                continue;
            }

            $pdo->prepare("INSERT INTO projects (job_order_no, project_name, province, region, location, assigned_team, remarks)
                VALUES (?, ?, ?, ?, ?, ?, ?)")
                ->execute([
                    $jobOrderNo,
                    $projectName,
                    trim((string)($row['province'] ?? '')),
                    trim((string)($row['region'] ?? '')),
                    trim((string)($row['location'] ?? '')),
                    trim((string)($row['assigned_team'] ?? $row['team'] ?? '')),
                    trim((string)($row['remarks'] ?? '')),
                ]);
            $existingJobOrders[$normalizedJo] = true;
            $existingProjectNames[$normalizedName] = true;
            $imported++;
        }

        if ($issues) {
            $shownIssues = array_slice($issues, 0, 5);
            $moreIssues = count($issues) > count($shownIssues) ? ' And ' . (count($issues) - count($shownIssues)) . ' more issue(s).' : '';
            flash('error', "Imported $imported project office(s); skipped $skipped. " . implode(' ', $shownIssues) . $moreIssues);
        } else {
            flash('success', "Imported $imported project office(s). Skipped $skipped duplicate or incomplete rows.");
        }
        redirect('?action=projects');
    }

    if ($postAction === 'cancel_import_project') {
        unset($_SESSION['pending_project_imports']);
        redirect('?action=projects');
    }

    if ($postAction === 'confirm_import_project') {
        $submitted = $_POST['projects'] ?? [];
        $_SESSION['pending_project_imports'] = array_map(static function ($row) {
            $row = is_array($row) ? $row : [];
            return [
                'job_order_no' => trim((string)($row['job_order_no'] ?? '')),
                'project_name' => trim((string)($row['project_name'] ?? '')),
                'province' => trim((string)($row['province'] ?? '')),
                'region' => trim((string)($row['region'] ?? '')),
                'location' => trim((string)($row['location'] ?? '')),
                'assigned_team' => trim((string)($row['assigned_team'] ?? '')),
                'remarks' => trim((string)($row['remarks'] ?? '')),
            ];
        }, is_array($submitted) ? $submitted : []);
        $submitted = is_array($submitted) ? $submitted : [];
        $projectsToSave = [];
        $seenJobOrders = [];
        $seenProjectNames = [];
        $existingJobOrders = [];
        $existingProjectNames = [];
        foreach ($pdo->query("SELECT job_order_no, project_name FROM projects") as $existingProject) {
            $existingJobOrders[normalize_job_order_no((string)$existingProject['job_order_no'])] = true;
            $existingProjectNames[normalize_project_name((string)$existingProject['project_name'])] = true;
        }
        foreach ($submitted as $rowIndex => $row) {
            $row = is_array($row) ? $row : [];
            $project = [
                trim((string)($row['job_order_no'] ?? '')),
                trim((string)($row['project_name'] ?? '')),
                trim((string)($row['province'] ?? '')),
                trim((string)($row['region'] ?? '')),
                trim((string)($row['location'] ?? '')),
                trim((string)($row['assigned_team'] ?? '')),
                trim((string)($row['remarks'] ?? '')),
            ];
            $projectsToSave[] = $project;
            $normalizedJo = normalize_job_order_no($project[0]);
            $normalizedName = normalize_project_name($project[1]);
            $error = '';
            $rowNumber = $rowIndex + 1;
            if ($project[0] === '' || $project[1] === '') {
                $missing = [];
                if ($project[0] === '') $missing[] = 'Job Order No.';
                if ($project[1] === '') $missing[] = 'Project Name';
                $error = 'Imported project ' . $rowNumber . ' is missing ' . implode(' and ', $missing) . '.';
            } elseif (isset($seenJobOrders[$normalizedJo])) {
                $error = 'Imported project ' . $rowNumber . ' has Job Order No. "' . $project[0] . '", which is duplicated in this file.';
            } elseif (isset($existingJobOrders[$normalizedJo])) {
                $error = 'Imported project ' . $rowNumber . ' has Job Order No. "' . $project[0] . '", which already exists in the project list.';
            } elseif (isset($seenProjectNames[$normalizedName])) {
                $error = 'Imported project ' . $rowNumber . ' has Project Name "' . $project[1] . '", which is duplicated in this file (case-insensitive match).';
            } elseif (isset($existingProjectNames[$normalizedName])) {
                $error = 'Imported project ' . $rowNumber . ' has Project Name "' . $project[1] . '", which already exists in the project list (case-insensitive match).';
            } else {
                $seenJobOrders[$normalizedJo] = true;
                $seenProjectNames[$normalizedName] = true;
            }
            if ($error !== '') {
                flash('error', $error);
                redirect('?action=projects&import_review=1');
            }
        }
        if (!$projectsToSave) {
            flash('error', 'There are no imported projects to save.');
            redirect('?action=projects');
        }
        $pdo->beginTransaction();
        try {
            $insert = $pdo->prepare("INSERT INTO projects (job_order_no, project_name, province, region, location, assigned_team, remarks) VALUES (?, ?, ?, ?, ?, ?, ?)");
            foreach ($projectsToSave as $project) {
                $insert->execute($project);
            }
            $pdo->commit();
        } catch (PDOException $e) {
            if ($pdo->inTransaction()) $pdo->rollBack();
            if ($e->getCode() === '23000' || stripos($e->getMessage(), 'uq_projects_job_order_no') !== false) {
                flash('error', 'A Job Order No. already exists. No projects were saved; review the numbers and try again.');
                redirect('?action=projects&import_review=1');
            }
            throw $e;
        }
        unset($_SESSION['pending_project_imports']);
        flash('success', 'All ' . count($projectsToSave) . ' imported project office(s) saved.');
        redirect('?action=projects');
    }

    if ($postAction === 'save_project') {
        $id = (int)($_POST['id'] ?? 0);
        $returnPage = max(1, (int)($_POST['return_page'] ?? 1));
        $data = [
            trim($_POST['job_order_no'] ?? ''),
            trim($_POST['project_name'] ?? ''),
            trim($_POST['province'] ?? ''),
            trim($_POST['region'] ?? ''),
            trim($_POST['location'] ?? ''),
            trim($_POST['assigned_team'] ?? ''),
            trim($_POST['remarks'] ?? '')
        ];

        if ($data[0] === '' || $data[1] === '') {
            $missing = [];
            if ($data[0] === '') $missing[] = 'Job Order No.';
            if ($data[1] === '') $missing[] = 'Project Name';
            flash('error', implode(' and ', $missing) . ' ' . (count($missing) === 1 ? 'is' : 'are') . ' required.');
            redirect($id > 0 ? '?action=projects&edit=' . $id . '&page=' . $returnPage : '?action=projects');
        }

        $duplicateQuery = $pdo->prepare("SELECT job_order_no, project_name FROM projects WHERE id <> ?");
        $duplicateQuery->execute([$id]);
        foreach ($duplicateQuery->fetchAll(PDO::FETCH_ASSOC) as $existingProject) {
            if (normalize_job_order_no((string)$existingProject['job_order_no']) === normalize_job_order_no($data[0])) {
                flash('error', 'Job Order No. "' . $data[0] . '" already exists (case-insensitive match).');
                redirect($id > 0 ? '?action=projects&edit=' . $id . '&page=' . $returnPage : '?action=projects');
            }
            if (normalize_project_name((string)$existingProject['project_name']) === normalize_project_name($data[1])) {
                flash('error', 'Project Name "' . $data[1] . '" already exists (case-insensitive match).');
                redirect($id > 0 ? '?action=projects&edit=' . $id . '&page=' . $returnPage : '?action=projects');
            }
        }

        try {
            if ($id > 0) {
                $stmt = $pdo->prepare("UPDATE projects
                    SET job_order_no=?, project_name=?, province=?, region=?, location=?, assigned_team=?, remarks=?
                    WHERE id=?");
                $stmt->execute([...$data, $id]);
                flash('success', 'Project updated.');
            } else {
                $stmt = $pdo->prepare("INSERT INTO projects
                    (job_order_no, project_name, province, region, location, assigned_team, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?)");
                $stmt->execute($data);
                flash('success', 'Project added.');
            }
        } catch (PDOException $e) {
            if (stripos($e->getMessage(), 'Duplicate entry') !== false || stripos($e->getMessage(), 'uq_projects_job_order_no') !== false) {
                flash('error', 'Job Order No. "' . $data[0] . '" already exists. Enter a unique number.');
                redirect($id > 0 ? '?action=projects&edit=' . $id . '&page=' . $returnPage : '?action=projects');
            }
            throw $e;
        }
        $projectRow = $pdo->query("SELECT province, project_name FROM projects WHERE id=" . $id)->fetch(PDO::FETCH_ASSOC);
        if ($id > 0) {
            $outages = $pdo->prepare("SELECT id FROM outages WHERE project_id=?");
            $outages->execute([$id]);
            $updateCode = $pdo->prepare("UPDATE outages SET outage_code=? WHERE id=?");
            foreach ($outages->fetchAll(PDO::FETCH_COLUMN) as $outageId) {
                $updateCode->execute([
                    outage_code((string)($projectRow['province'] ?? ''), (int)$outageId, (string)($projectRow['project_name'] ?? '')),
                    (int)$outageId
                ]);
            }
        }
        redirect('?action=projects&page=' . $returnPage . ($id > 0 ? '#project-row-' . $id : ''));
    }

    if ($postAction === 'delete_project') {
        $id = (int)($_POST['id'] ?? 0);
        $pdo->prepare("DELETE FROM projects WHERE id=?")->execute([$id]);
        flash('success', 'Project deleted.');
        redirect('?action=projects');
    }
