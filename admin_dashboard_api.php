
<?php
session_start();
header("Content-Type: application/json");

require_once "config/database.php";

if (
    !isset($_SESSION["user_id"]) ||
    $_SESSION["role"] !== "admin"
) {
    http_response_code(401);
    echo json_encode([
        "success" => false,
        "message" => "Admin login required"
    ]);
    exit;
}

try {
    $today = date("Y-m-d");

    // All employees
    $empStmt = $pdo->query("
        SELECT id, employee_code, name, email, phone,
               position, department, joining_date
        FROM employees
        ORDER BY id DESC
    ");

    $employees = $empStmt->fetchAll();

    // Today's attendance
    $todayStmt = $pdo->prepare("
        SELECT
            a.id,
            a.employee_id,
            a.attendance_date,
            a.check_in,
            a.check_out,
            a.status,
            a.total_hours,
            e.employee_code
        FROM attendance a
        INNER JOIN employees e ON a.employee_id = e.id
        WHERE a.attendance_date = ?
        ORDER BY a.id DESC
    ");

    $todayStmt->execute([$today]);
    $todayAttendance = $todayStmt->fetchAll();

    // Latest 8 attendance records
    $recentStmt = $pdo->query("
        SELECT
            a.id,
            a.employee_id,
            a.attendance_date,
            a.check_in,
            a.check_out,
            a.status,
            a.total_hours,
            e.employee_code,
            e.name,
            e.position,
            e.department
        FROM attendance a
        INNER JOIN employees e ON a.employee_id = e.id
        ORDER BY a.attendance_date DESC, a.id DESC
        LIMIT 8
    ");

    $recentActivity = $recentStmt->fetchAll();

    echo json_encode([
        "success" => true,
        "date" => $today,
        "employees" => $employees,
        "todayAttendance" => $todayAttendance,
        "recentActivity" => $recentActivity
    ]);

} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode([
        "success" => false,
        "message" => "Database error: " . $e->getMessage()
    ]);
}
?>
