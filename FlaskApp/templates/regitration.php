<?php
session_start();
error_reporting(E_ALL);
ini_set('display_errors', 1);

// Initialize error messages
$emailError = $usernameError = $passwordError = $registrationError = "";

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Get user input (email, username, and password)
    $email = $_POST['email'];
    $username = $_POST['username'];
    $password = $_POST['password'];
    
    // Hash the password
    $hashedPassword = password_hash($password, PASSWORD_DEFAULT);
    
    // Store the username and hashed password in your database
    // Replace the database connection details and SQL query with your own
    $dbHost = 'localhost';
    $dbName = 'tracker';
    $dbUser = 'test';
    $dbPass = 'test';
    $date = $time = date('Y-m-d H:i:s');
    
    try {
        $conn = new PDO("mysql:host=$dbHost;dbname=$dbName", $dbUser, $dbPass);
        $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

        // Check for empty email, username, and password
        if (empty($email)) {
            $emailError = "Email is required.";
        }
        if (empty($username)) {
            $usernameError = "Username is required.";
        }
        if (empty($password)) {
            $passwordError = "Password is required.";
        }

        $checkSql = "SELECT * FROM registration WHERE email = :email OR username = :username";
        $checkStmt = $conn->prepare($checkSql);
        $checkStmt->bindParam(":email", $email, PDO::PARAM_STR);
        $checkStmt->bindParam(":username", $username, PDO::PARAM_STR);
        $checkStmt->execute();

        if ($checkStmt->rowCount() > 0) {
            $registration2Error = "Email or Username already in use.";
        } else if (empty($emailError) && empty($usernameError) && empty($passwordError)) {
            $sql = "INSERT INTO registration (email, username, password_hash, time, date, userType) VALUES (:email, :username, :password, '$time', '$date', 0)";
            $stmt = $conn->prepare($sql);
            $stmt->bindParam(":email", $email, PDO::PARAM_STR);
            $stmt->bindParam(":username", $username, PDO::PARAM_STR);
            $stmt->bindParam(":password", $hashedPassword, PDO::PARAM_STR);
            $stmt->execute();
            $_SESSION['username'] = $username;
            header("Location: http://127.0.0.1:5000"); // Redirect to a protected area
        } else {
            $registrationError = "Registration failed. Please fix the errors.";
        }
    } catch (PDOException $e) {
        echo "Error: " . $e->getMessage();
    }
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Registerartion Form</title>
<link rel="shortcut icon" type="image/x-icon" href="https://github.com/Shumokh1/Research-Tracker/blob/main/pictures/MagicEraser_231212_123200.png?raw=true">
<!-- background-color: rgb(147, 142, 214); -->
        <style>
        body {
            font-family: Arial, sans-serif;
            background-image: url("https://github.com/Shumokh1/Research-Tracker/blob/main/pictures/MagicEraser_231212_123200.png?raw=true");
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
        }
        .form-container {
            background: white;
            padding: 20px;
            margin: 10px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
        }
        input[type="username"], input[type="password"], input[type="email"]{
            width: calc(100% - 20px);
            margin-bottom: 15px;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        .btn {
            background-color: #4CAF50;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            width: 100%;
        }
        .btn:hover {
            opacity: 0.8;
        }
        .form-footer {
            margin-top: 10px;
            text-align: center;
        }

        body {
            background-color: rgb(255, 255, 255);
            color: rgb(133, 87, 133);
            background-size: cover;
            backdrop-filter: blur(5px); /* Apply blur effect */
            background-attachment: fixed;
        }

        nav{
            background-color: rgb(236, 268, 211);
        }

        header, footer {
            background-color: rgb(236, 168, 211);
            color: rgb(255, 240, 255);
        }

        .btn-primary {
            background-color: rgb(147, 142, 214);
            color: rgb(255, 255, 255);
            border: none;
        }

        .btn-primary:hover {
            background-color: rgb(85, 79, 156));
            color: rgb(255, 255, 255);
        }

        /* Additional styling as needed */


        /* Bootstrap Centering and Responsiveness */
        .centered-container {
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh; /* Full viewport height */
            
        }

        /* Existing styles */
        /* ... */

        /* Additional responsive adjustments */
        @media (max-width: 768px) {
            /* Adjustments for smaller screens */
        }

        .help-content {
            display: none;
            position: absolute;
            background-color: #f9f9f9;
            box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
            padding: 12px;
            border-radius: 5px;
            z-index: 1;
            bottom: 5%; /* Position the div above the button */
            left: 50%; /* Start from the middle of the button */
            transform: translateX(-50%); /* Center the div relative to the button */
            
        }
        </style>
</head>
<body>

  <div class="row">
    <div class="col-6">
      <div class="form-container">
        <h2>Registeration</h2>
          <form method="post" action="<?php echo $_SERVER['PHP_SELF']; ?>">
              <?php
                if (!empty($registrationError)) {
                    echo "<p style='color: red;'>$registrationError</p>";
                }else if (!empty($registration2Error))
                {
                    echo "<p style='color: red;'>$registration2Error</p>";
                }
            ?>

              <input name="email" type="email" placeholder="Enter your email" required>
              <span style="color: red;"><?php echo $emailError; ?></span>
              <input name="username" type="username" placeholder="Enter your username" required>
              <span style="color: red;"><?php echo $usernameError; ?></span>
              <input name="password" type="password" placeholder="Enter your password" required>
              <button type="submit" class="btn btn-primary" name="register">register</button>
              <div class="form-footer">
              Have an exiting account? <a onclick="login()" href="#">SignIn</a>
                <br><br>
                <center><a href="#" onmouseover="showHelp()" onmouseout="hideHelp()">Help</a>
                <div id="helpDiv" class="help-content">
                    <p>This website assists educators and students in locating studies published by particular researchers.</p>
                </div></center>
        </form>
      </div>
    </div>

  </div>

</body>

<script>
    function login()
    {
        window.location.href = 'login.php';
        return false;
    }

    function showHelp() {
        document.getElementById('helpDiv').style.display = 'block';
    }

    function hideHelp() {
        document.getElementById('helpDiv').style.display = 'none';
    }
</script>
</html>