<?php
$servername = "localhost";
$db_username = "root";
$db_password = "echo";

$conn = new mysqli($servername, $db_username, $db_password, "saleproject");
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
} 

?>