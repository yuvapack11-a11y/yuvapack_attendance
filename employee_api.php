
<?php
session_start();
header("Content-Type: application/json");

require_once "config/database.php";

function response($success, $message, $data = null) {
    echo json_encode([
        "success" => $success,
        "message" => $message,
        "data" => $data
    ]);
    exit;
}

// Admin access only
if (
    !isset($_SESSION["user_id"]) ||
    !isset($_SESSION["role"]) ||
    $_SESSION["role"] !== "admin"
) {
    http_response_code(401);
    response(false, "Unauthorized access");
}

$method = $_SERVER["REQUEST_METHOD"];
$data = json_decode(file_get_contents("php://input"), true) ?? [];

// ==================================================
// GET — Load active employees
// ==================================================
if ($method === "GET") {
    try {
        $stmt = $pdo->query("
            SELECT
                e.id,
                e.employee_code,
                e.name,
                e.email,
                e.phone,
                e.position,
                e.department,
                e.joining_date,
                e.user_id,
                u.username
            FROM employees e
            LEFT JOIN users u ON e.user_id = u.id
            WHERE e.is_active = 1
            ORDER BY e.id DESC
        ");

        response(true, "Employees loaded", $stmt->fetchAll());

    } catch (PDOException $e) {
        http_response_code(500);
        response(false, "Database error");
    }
}

// ==================================================
// POST — Add new employee
// ==================================================
if ($method === "POST") {
    $code = trim($data["employee_code"] ?? "");
    $name = trim($data["name"] ?? "");
    $email = trim($data["email"] ?? "");
    $phone = trim($data["phone"] ?? "");
    $position = trim($data["position"] ?? "");
    $department = trim($data["department"] ?? "");
    $joiningDate = $data["joining_date"] ?? "";
    $password = $data["password"] ?? "";

    if (!$code || !$name || !$password) {
        http_response_code(400);
        response(false, "Employee code, name and password are required");
    }

    try {
        $pdo->beginTransaction();

        $check = $pdo->prepare(
            "SELECT id FROM employees WHERE employee_code = ?"
        );
        $check->execute([$code]);

        if ($check->fetch()) {
            $pdo->rollBack();
            http_response_code(409);
            response(false, "Employee code already exists");
        }

        $checkUser = $pdo->prepare(
            "SELECT id FROM users WHERE username = ?"
        );
        $checkUser->execute([$code]);

        if ($checkUser->fetch()) {
            $pdo->rollBack();
            http_response_code(409);
            response(false, "Username already exists");
        }

        $hashedPassword = password_hash($password, PASSWORD_DEFAULT);

        $userStmt = $pdo->prepare("
            INSERT INTO users (username, password, role)
            VALUES (?, ?, 'employee')
        ");
        $userStmt->execute([$code, $hashedPassword]);

        $userId = $pdo->lastInsertId();

        $employeeStmt = $pdo->prepare("
            INSERT INTO employees
            (employee_code, name, email, phone, position,
             department, joining_date, user_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ");

        $employeeStmt->execute([
            $code,
            $name,
            $email ?: null,
            $phone ?: null,
            $position ?: null,
            $department ?: null,
            $joiningDate ?: null,
            $userId
        ]);

        $employeeId = $pdo->lastInsertId();

        $pdo->commit();

        response(true, "Employee added successfully", [
            "id" => $employeeId,
            "employee_code" => $code
        ]);

    } catch (Throwable $e) {
        if ($pdo->inTransaction()) {
            $pdo->rollBack();
        }

        http_response_code(500);
        response(false, "Could not add employee");
    }
}

// ==================================================
// PUT — Edit employee
// ==================================================
if ($method === "PUT") {
    $code = trim($data["employee_code"] ?? "");
    $name = trim($data["name"] ?? "");
    $email = trim($data["email"] ?? "");
    $phone = trim($data["phone"] ?? "");
    $position = trim($data["position"] ?? "");
    $department = trim($data["department"] ?? "");
    $joiningDate = $data["joining_date"] ?? "";
    $password = $data["password"] ?? "";

    if (!$code || !$name) {
        http_response_code(400);
        response(false, "Employee code and name are required");
    }

    try {
        $pdo->beginTransaction();

        $find = $pdo->prepare("
            SELECT id, user_id
            FROM employees
            WHERE employee_code = ? AND is_active = 1
        ");
        $find->execute([$code]);
        $employee = $find->fetch();

        if (!$employee) {
            $pdo->rollBack();
            http_response_code(404);
            response(false, "Employee not found");
        }

        $update = $pdo->prepare("
            UPDATE employees
            SET name = ?,
                email = ?,
                phone = ?,
                position = ?,
                department = ?,
                joining_date = ?
            WHERE employee_code = ? AND is_active = 1
        ");

        $update->execute([
            $name,
            $email ?: null,
            $phone ?: null,
            $position ?: null,
            $department ?: null,
            $joiningDate ?: null,
            $code
        ]);

        // Password update is optional
        if ($password !== "" && !empty($employee["user_id"])) {
            $hashedPassword = password_hash(
                $password,
                PASSWORD_DEFAULT
            );

            $updatePassword = $pdo->prepare("
                UPDATE users
                SET password = ?
                WHERE id = ? AND role = 'employee'
            ");

            $updatePassword->execute([
                $hashedPassword,
                $employee["user_id"]
            ]);
        }

        $pdo->commit();

        response(true, "Employee updated successfully");

    } catch (Throwable $e) {
        if ($pdo->inTransaction()) {
            $pdo->rollBack();
        }

        http_response_code(500);
        response(false, "Could not update employee");
    }
}

// ==================================================
// DELETE — Soft delete employee
// ==================================================
if ($method === "DELETE") {
    $code = trim($data["employee_code"] ?? "");

    if (!$code) {
        http_response_code(400);
        response(false, "Employee code is required");
    }

    try {
        $stmt = $pdo->prepare("
            UPDATE employees
            SET is_active = 0
            WHERE employee_code = ? AND is_active = 1
        ");

        $stmt->execute([$code]);

        if ($stmt->rowCount() === 0) {
            http_response_code(404);
            response(false, "Employee not found or already deleted");
        }

        response(true, "Employee deleted successfully");

    } catch (PDOException $e) {
        http_response_code(500);
        response(false, "Could not delete employee");
    }
}

http_response_code(405);
response(false, "Method not allowed");
