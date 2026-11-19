
<?php
session_start();
header("Content-Type: application/json");

require_once "config/database.php";

if ($_SERVER["REQUEST_METHOD"] !== "POST") {
    http_response_code(405);
    echo json_encode([
        "success" => false,
        "message" => "POST method required"
    ]);
    exit;
}

$data = json_decode(file_get_contents("php://input"), true);

$id = trim($data["employee_code"] ?? "");
$password = $data["password"] ?? "";

if ($id === "" || $password === "") {
    http_response_code(400);
    echo json_encode([
        "success" => false,
        "message" => "Employee ID and password required"
    ]);
    exit;
}

$stmt = $pdo->prepare("
    SELECT
        e.id,
        e.employee_code,
        e.name,
        e.email,
        e.phone,
        e.position,
        e.department,
        e.joining_date,
        u.id AS user_id,
        u.password,
        u.role
    FROM employees e
    INNER JOIN users u ON e.user_id = u.id
    WHERE e.employee_code = ?
    LIMIT 1
");

$stmt->execute([$id]);
$emp = $stmt->fetch();

if (
    !$emp ||
    $emp["role"] !== "employee" ||
    !password_verify($password, $emp["password"])
) {
    http_response_code(401);
    echo json_encode([
        "success" => false,
        "message" => "Invalid employee ID or password"
    ]);
    exit;
}

session_regenerate_id(true);

$_SESSION["user_id"] = $emp["user_id"];
$_SESSION["employee_id"] = $emp["id"];
$_SESSION["employee_code"] = $emp["employee_code"];
$_SESSION["role"] = "employee";

unset($emp["password"]);

echo json_encode([
    "success" => true,
    "message" => "Employee login successful",
    "employee" => $emp
]);
?>
