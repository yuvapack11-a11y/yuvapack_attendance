
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

$username = trim($data["username"] ?? "");
$password = $data["password"] ?? "";

if ($username === "" || $password === "") {
    http_response_code(400);
    echo json_encode([
        "success" => false,
        "message" => "Username and password required"
    ]);
    exit;
}

$stmt = $pdo->prepare(
    "SELECT id, username, password, role
     FROM users
     WHERE username = ?
     LIMIT 1"
);

$stmt->execute([$username]);
$user = $stmt->fetch();

if (!$user || !password_verify($password, $user["password"])) {
    http_response_code(401);
    echo json_encode([
        "success" => false,
        "message" => "Invalid username or password"
    ]);
    exit;
}

session_regenerate_id(true);

$_SESSION["user_id"] = $user["id"];
$_SESSION["username"] = $user["username"];
$_SESSION["role"] = $user["role"];

echo json_encode([
    "success" => true,
    "message" => "Login successful",
    "user" => [
        "id" => $user["id"],
        "username" => $user["username"],
        "role" => $user["role"]
    ]
]);
?>
