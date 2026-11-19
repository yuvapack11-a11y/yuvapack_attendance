
<?php
session_start();
header("Content-Type: application/json");

require_once "config/database.php";

date_default_timezone_set("Asia/Kolkata");

function sendResponse($success, $message, $data = null) {
    echo json_encode([
        "success" => $success,
        "message" => $message,
        "data" => $data
    ]);
    exit;
}

// Employee session check
if (
    !isset($_SESSION["user_id"]) ||
    !isset($_SESSION["employee_id"]) ||
    $_SESSION["role"] !== "employee"
) {
    http_response_code(401);
    sendResponse(false, "Employee login required");
}

$employeeId = (int) $_SESSION["employee_id"];
$today = date("Y-m-d");

if ($_SERVER["REQUEST_METHOD"] === "GET") {
    $stmt = $pdo->prepare("
        SELECT id, employee_id, attendance_date,
               check_in, check_out, status, total_hours
        FROM attendance
        WHERE employee_id = ? AND attendance_date = ?
        LIMIT 1
    ");
    $stmt->execute([$employeeId, $today]);

    sendResponse(true, "Attendance loaded", $stmt->fetch() ?: null);
}

if ($_SERVER["REQUEST_METHOD"] !== "POST") {
    http_response_code(405);
    sendResponse(false, "Method not allowed");
}

$data = json_decode(file_get_contents("php://input"), true);
$action = $data["action"] ?? "";

try {
    if ($action === "check_in") {
        $now = date("H:i:s");

        // Prevent duplicate check-in for today
        $stmt = $pdo->prepare("
            SELECT id, check_in
            FROM attendance
            WHERE employee_id = ? AND attendance_date = ?
            LIMIT 1
        ");
        $stmt->execute([$employeeId, $today]);
        $record = $stmt->fetch();

        if ($record && $record["check_in"]) {
            http_response_code(409);
            sendResponse(false, "Already checked in at " . $record["check_in"]);
        }

        // Office start 09:00 AM; grace period until 09:15 AM
        $status = ($now > "09:15:00") ? "Late" : "Present";

        if ($record) {
            $stmt = $pdo->prepare("
                UPDATE attendance
                SET check_in = ?, status = ?
                WHERE id = ?
            ");
            $stmt->execute([$now, $status, $record["id"]]);
        } else {
            $stmt = $pdo->prepare("
                INSERT INTO attendance
                (employee_id, attendance_date, check_in, status, total_hours)
                VALUES (?, ?, ?, ?, 0)
            ");
            $stmt->execute([$employeeId, $today, $now, $status]);
        }

        sendResponse(true, "Check-in successful", [
            "date" => $today,
            "check_in" => $now,
            "status" => $status
        ]);
    }

    if ($action === "check_out") {
        $stmt = $pdo->prepare("
            SELECT id, check_in, check_out
            FROM attendance
            WHERE employee_id = ? AND attendance_date = ?
            LIMIT 1
        ");
        $stmt->execute([$employeeId, $today]);
        $record = $stmt->fetch();

        if (!$record || !$record["check_in"]) {
            http_response_code(400);
            sendResponse(false, "You must check in first");
        }

        if ($record["check_out"]) {
            http_response_code(409);
            sendResponse(false, "Already checked out at " . $record["check_out"]);
        }

        $checkOut = date("H:i:s");

        $start = strtotime($today . " " . $record["check_in"]);
        $end = strtotime($today . " " . $checkOut);

        if ($end < $start) {
            http_response_code(400);
            sendResponse(false, "Check-out time cannot be earlier than check-in time");
        }

        $hours = round(($end - $start) / 3600, 2);

        $stmt = $pdo->prepare("
            UPDATE attendance
            SET check_out = ?, total_hours = ?
            WHERE id = ?
        ");
        $stmt->execute([$checkOut, $hours, $record["id"]]);

        sendResponse(true, "Check-out successful", [
            "date" => $today,
            "check_in" => $record["check_in"],
            "check_out" => $checkOut,
            "total_hours" => $hours
        ]);
    }

    http_response_code(400);
    sendResponse(false, "Invalid attendance action");

} catch (PDOException $e) {
    http_response_code(500);
    sendResponse(false, "Database error: " . $e->getMessage());
}
