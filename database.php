
<?php

$host = "localhost";
$dbname = "yuvapack_attendance";
$username = "root";
$password = "";

date_default_timezone_set("Asia/Kolkata");

try {
    $pdo = new PDO(
        "mysql:host=$host;dbname=$dbname;charset=utf8mb4",
        $username,
        $password
    );

    $pdo->setAttribute(
        PDO::ATTR_ERRMODE,
        PDO::ERRMODE_EXCEPTION
    );

    $pdo->setAttribute(
        PDO::ATTR_DEFAULT_FETCH_MODE,
        PDO::FETCH_ASSOC
    );

} catch (PDOException $e) {
    die("Database connection failed: " . $e->getMessage());
}

?>
