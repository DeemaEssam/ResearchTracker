<?php
ini_set('display_errors', 1);
session_start();

// Initialize error messages
$usernameError = $passwordError = $loginError = "";

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Get user input (username and password)
    $username = $_POST['username'];
    $password = $_POST['password'];

    // Replace the database connection details with your own
    $dbHost = 'localhost';
    $dbName = 'tracker';
    $dbUser = 'test';
    $dbPass = 'test';
    $date = $time = date('Y-m-d H:i:s');

    try {
        $conn = new PDO("mysql:host=$dbHost;dbname=$dbName", $dbUser, $dbPass);
        $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

        // Check for empty username and password
        if (empty($username)) {
            $usernameError = "Username is required.";
        }
        if (empty($password)) {
            $passwordError = "Password is required.";
        }

        // If there are no validation errors, proceed with login
        if (empty($usernameError) && empty($passwordError)) {
            // Fetch the hashed password from the database based on the provided username
            $sql = "SELECT password_hash FROM registration WHERE username = :username";
            $stmt = $conn->prepare($sql);
            $stmt->bindParam(":username", $username, PDO::PARAM_STR);
            $stmt->execute();
            $row = $stmt->fetch(PDO::FETCH_ASSOC);

            if ($row) {
                $hashedPasswordFromDatabase = $row['password_hash'];

                // Verify the provided password with the hashed password from the database
                if (password_verify($password, $hashedPasswordFromDatabase)) {
                    // Passwords match, user is authenticated
                    $sql2 = "INSERT INTO login (time, date, username) VALUES ('$time', '$date', :username)";
                    $stmt2 = $conn->prepare($sql2);
                    $stmt2->bindParam(":username", $username, PDO::PARAM_STR);
                    $stmt2->execute();
                    $_SESSION['username'] = $username; // Store username in the session
                    header("Location:  http://127.0.0.1:5000");
                    exit();
                } else {
                    // Passwords do not match, display an error message
                    $loginError = "Invalid username or password.";
                }
            } else {
                // User does not exist, display an error message
                $loginError = "Invalid username or password.";
            }
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
        <title>Login Form</title>
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
        input[type="username"], input[type="password"]{
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

        nave{
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

    <div class="form-container">
        <h2>Login</h2>
        <form method="post" action="<?php echo $_SERVER['PHP_SELF']; ?>">
            <?php
                if (!empty($loginError)) {
                    echo "<p style='color: red;'>$loginError</p>";
                }
            ?>
            <input type="username" name="username" required id="username" placeholder="Enter your username"><br>
            <span style="color: red;"><?php echo $usernameError; ?></span>
            <input type="password" name="password" required id="password" placeholder="Enter your password"><br>
            <span style="color: red;"><?php echo $passwordError; ?></span>

            <!-- <a href="#">Forgot password?</a><br> -->
            <button type="submit" class="btn btn-primary mt-3" name="login">Login</button><br>
            Don't have an account? <a Onclick="reg()" href="#">SignUp</a>
            <br><br>
            <center><a href="#" onmouseover="showHelp()" onmouseout="hideHelp()">Help</a>
            <div id="helpDiv" class="help-content">
                <p>This website assists educators and students in locating studies published by particular researchers.</p>
            </div></center>
        </form>
    </div>

</body>

<script>

  function reg()
  {
     window.location.href = 'regitration.php';
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
