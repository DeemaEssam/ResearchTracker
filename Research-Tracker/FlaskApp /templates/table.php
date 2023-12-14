<?php

session_start();

  ini_set('display_errors', 1);

  $connection = mysqli_connect("localhost","test","test");
  $db = mysqli_select_db($connection, 'tracker');

  $sql = "select * from login";
  $result = mysqli_query($connection,$sql);

// Check if the user is not logged in, redirect to the login page
if (isset($_SESSION['username'])) {
// The rest of your protected page's code goes here
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitor Page</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="shortcut icon" type="image/x-icon" href="https://github.com/Shumokh1/Research-Tracker/blob/main/pictures/MagicEraser_231212_123200.png?raw=true">
<!-- background-color: rgb(255, 240, 251); -->

        <style>
        body {
            font-family: Arial, sans-serif;
            background-image: url("https://github.com/Shumokh1/Research-Tracker/blob/main/pictures/MagicEraser_231212_123200.png?raw=true");
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
            background-color: rgb(255, 255, 255);
            color: rgb(133, 87, 133);
            background-size: cover;
            backdrop-filter: blur(5px); /* Apply blur effect */
            background-attachment: fixed;
        }

        nav{
            background-color: rgb(236, 268, 211);
        }

        header {
            background-color: rgb(236, 168, 211);
            color: rgb(255, 240, 255);
        }

        .centered-container {
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh; /* Full viewport height */
        }

        .no-results-message {
          text-align: center;
          margin-top: 20px;
          padding: 10px;
          background-color: #ffcccc; /* Light red background */
          color: #cc0000; /* Darker red text */
          border: 1px solid #cc0000;
          border-radius: 5px;
      }
        </style>
</head>
<body>

  <!-- <header>
    <nav class="navbar navbar-expand-lg bg-body-tertiary">
      <div class="container-fluid">
        <a class="navbar-brand" Onclick="home()">Home</a>
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarText"
          aria-controls="navbarText" aria-expanded="false" aria-label="Toggle navigation">
          <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarText">
          <ul class="navbar-nav me-auto mb-2 mb-lg-0">
            <li class="nav-item">
              <a class="nav-link active" aria-current="page" Onclick="mon()">Monitor</a>
            </li>
          </ul>
          <span class="navbar-text">
            Logout
          </span>
        </div>
      </div>
    </nav>
  </header> -->

    <div class="container">
        <div>
            <table class="table table-striped mt-3">
                <thead>
                    <tr><?php if (isset($_SESSION['username'])) {
                        $mysqli = new mysqli("localhost", "test", "test", "tracker");
                        $result2 = $mysqli->query("SELECT userType FROM registration WHERE userType=1 AND username='$_SESSION[username]'");
                        if ($result2->num_rows == 1)  {?>
                            <th class="table-light text-center" scope="col">Time</th>
                            <th class="table-light text-center" scope="col">Date</th>
                            <th class="table-light text-center" scope="col">username</th>
                        <?php }else{?><br>
                            <div class="no-results-message">
                                Access to this page is restricted to privileged users.
                            </div>
                        <?php }}?>
                    </tr>
                </thead>
                <tbody>
                    <?php 
                    $i = 0;
                    while($rows = mysqli_fetch_assoc($result)){
                    $i++;
                    ?>
                    <tr>
                        <?php if (isset($_SESSION['username'])) {
                            $mysqli = new mysqli("localhost", "test", "test", "tracker");
                            $result2 = $mysqli->query("SELECT userType FROM registration WHERE userType=1 AND username='$_SESSION[username]'");
                            if ($result2->num_rows == 1)  {?>
                                <th class="table-light text-center" scope="col"><?php echo $rows['time']; ?></th>
                                <th class="table-light text-center" scope="col"><?php echo $rows['date']; ?></th>
                                <th class="table-light text-center" scope="col"><?php echo $rows['username']; ?></th>
                            <?php } ?>
                                
                            <?php 
                        } else{ ?>
                            <th>Access to this page is restricted to privileged users.</th>
                        <?php } ?>
                    </tr>
                </tbody>
                <?php } ?>
            </table>
        </div>
    </div>
    
    </body>

    <script>
        function home()
            {
                window.location.href = 'http://127.0.0.1:5000';
                return false;
            }
    </script>
</html>

<?php
}else{
    header("Location: login.php"); // Replace "login.php" with the actual login page URL
    exit();
}
?>