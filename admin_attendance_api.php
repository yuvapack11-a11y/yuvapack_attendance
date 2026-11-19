
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

if ($_SERVER["REQUEST_METHOD"] !== "GET") {
    http_response_code(405);
    echo json_encode([
        "success" => false,
        "message" => "GET method required"
    ]);
    exit;
}

try {
    $stmt = $pdo->query("
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
            e.department,
            e.phone
        FROM attendance a
        INNER JOIN employees e
            ON a.employee_id = e.id
        ORDER BY a.attendance_date DESC, a.id DESC
    ");

    $records = $stmt->fetchAll();

    echo json_encode([
        "success" => true,
        "message" => "Attendance records loaded",
        "data" => $records
    ]);

} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode([
        "success" => false,
        "message" => "Database error: " . $e->getMessage()
    ]);
}
?>
