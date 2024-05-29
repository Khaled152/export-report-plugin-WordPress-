<?php
/*
Plugin Name: Export Courses to Excel
Description: Plugin to export courses data to Excel using PHPSpreadsheet
Version: 1.0
Author: Your Name
*/

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once plugin_dir_path( __FILE__ ) . 'vendor/autoload.php';

// Function to execute SQL query and export results to Excel
function export_courses_to_excel() {
    global $wpdb;
    $table_name = $wpdb->prefix . 'courses';
    // Check user capabilities
    if ( ! current_user_can( 'manage_options' ) ) {
        return;
    }

    // SQL query
    $query = ' SELECT _q_.`base.User_Login` AS `HR Code`,
            _q_.`base.User_Display_Name` AS `Employee Name`,
            _q_.`base.User_Email` AS `User Email`,
            _q_.`base.Enrollment_Date` AS `Enrollment Date`,
            comp.completion_date AS `complete date`,
            progress.percentage as "progress",
            complete_status.completed As `completion status`,
            res.mark as "Pre assessment score",
            assessment.mark2 as "post assessment score",
            assessment.status as "status",
            _q_.`base.Course_ID` AS `Course ID`,
            _q_.`base.Course_Name` AS `Course Name`,
            _q_.`base.User_ID` AS `User ID`,
            _q_.`base.User_Login` AS `User Login`,
            `is`.`Completion_ID` AS `Completion ID`,
            `is`.`Certificate_ID` AS `Certificate ID`,
            _q_.`base.Enrollment_ID` AS `Enrollment_ID`,
            retry.comment_count AS `Course retake`
        FROM 
            (
                SELECT base.`Enrollment_ID` AS `base.Enrollment_ID`, 
                    base.`Enrollment_Date` AS `base.Enrollment_Date`, 
                    base.`Course_ID` AS `base.Course_ID`, 
                    base.`Course_Name` AS `base.Course_Name`, 
                    base.`User_ID` AS `base.User_ID`, 
                    base.`User_Login` AS `base.User_Login`, 
                    base.`User_Email` AS `base.User_Email`, 
                    base.`User_Display_Name` AS `base.User_Display_Name`
                FROM
                    (
                        SELECT `Enrollment_ID` as `Enrollment_ID`, 
                            `Enrollment_Date` as `Enrollment_Date`, 
                            `Course_ID` as `Course_ID`, 
                            `Course_Name` as `Course_Name`, 
                            `User_ID` as `User_ID`, 
                            `User_Login` as `User_Login`, 
                            `User_Email` as `User_Email`, 
                            `User_Display_Name` as `User_Display_Name`
                        FROM 
                            (
                                SELECT `wp`.`ID` AS `Enrollment_ID`,
                                    `wp`.`post_date` AS `Enrollment_Date`,
                                    `wp`.`post_parent` AS `Course_ID`,
                                    `wp1`.`post_title` AS `Course_Name`,
                                    `wu1`.`ID` AS `User_ID`,
                                    `wu1`.`user_login` AS `User_Login`,
                                    `wu1`.`user_email` AS `User_Email`,
                                    `wu1`.`display_name` AS `User_Display_Name`
                                FROM `wp_posts` AS `wp`
                                LEFT JOIN `wp_posts` AS `wp1` ON `wp`.`post_parent` = `wp1`.`ID`
                                LEFT JOIN `wp_users` AS `wu1` ON `wp`.`post_author` = `wu1`.`ID`
                                WHERE (((`wp`.`post_type`) = "tutor_enrolled"))
                                AND (((`wp`.`post_date`) >= timestamp(MAKEDATE(YEAR(CURRENT_DATE()),1)) AND ((`wp`.`post_date`) < (timestamp(MAKEDATE(YEAR(CURRENT_DATE()),1)) + interval 1 year))))
                                AND (((`wp`.`post_parent`) = 7619))
                            ) AS q
                    ) AS `base`
            ) AS _q_
        LEFT JOIN 
            (
                SELECT `Completion_ID` as `Completion_ID`, 
                    `Completion_Date` as `Completion_Date`, 
                    `Certificate_ID` as `Certificate_ID`, 
                    `Course_ID` as `Course_ID`, 
                    `Course_Name` as `Course_Name`, 
                    `User_ID` as `User_ID`, 
                    `User_Login` as `User_Login`, 
                    `User_Email` as `User_Email`, 
                    `User_Display_Name` as `User_Display_Name`
                FROM 
                    (
                        SELECT `x`.`comment_ID` AS `Completion_ID`,
                            `x`.`comment_date` AS `Completion_Date`,
                            `x`.`comment_content` AS `Certificate_ID`,
                            `wp`.`ID` AS `Course_ID`,
                            `wp`.`post_title` AS `Course_Name`,
                            `wu`.`ID` AS `User_ID`,
                            `wu`.`user_login` AS `User_Login`,
                            `wu`.`user_email` AS `User_Email`,
                            `wu`.`display_name` AS `User_Display_Name`
                        FROM `wp_comments` AS `x`
                        LEFT JOIN `wp_posts` AS `wp` ON `x`.`comment_post_ID` = `wp`.`ID`
                        LEFT JOIN `wp_users` AS `wu` ON `x`.`comment_author` = `wu`.`ID`
                        WHERE (((`x`.`comment_type`) = "course_completed"))
                        AND (((`wp`.`ID`) = 7619))
                    ) AS q
            ) as `is` ON _q_.`base.User_ID` = `is`.`User_ID`
        LEFT JOIN 
            (
                SELECT quiz_id , wp_tutor_quiz_attempt_answers .user_id,  concat(round(sum(achieved_mark) / sum(question_mark) * 100 , 2 ) , "%")  as mark 
                FROM wp_users 
                join wp_tutor_quiz_attempt_answers on wp_users.id = wp_tutor_quiz_attempt_answers.user_id 
                WHERE quiz_id = 8044 
                group by quiz_id , wp_tutor_quiz_attempt_answers .user_id
            ) as `res` ON _q_.`base.User_ID` = `res`.`user_ID`
        LEFT JOIN 
            (
                SELECT quiz_id, user_id, mark2,
                    CASE
                        WHEN mark2 <= 60 THEN "failed"
                        ELSE "success"
                    END as status
                FROM (
                    SELECT quiz_id , wp_tutor_quiz_attempt_answers.user_id, 
                            concat(round(sum(achieved_mark) / sum(question_mark) * 100 , 2 ) , "%")  as mark2
                    FROM wp_users
                    JOIN wp_tutor_quiz_attempt_answers on wp_users.id = wp_tutor_quiz_attempt_answers.user_id
                    WHERE quiz_id = "9448"
                    GROUP BY quiz_id , wp_tutor_quiz_attempt_answers.user_id
                ) as subquery 
            ) as `assessment` ON _q_.`base.User_ID` = `assessment`.`user_ID`
        LEFT JOIN 
            (
                WITH quiz_counts AS (
                    SELECT user_id, 
                        COUNT(DISTINCT quiz_id) AS total_quizzes 
                    FROM wp_tutor_quiz_attempts 
                    WHERE course_id = 7619 
                    GROUP BY user_id
                ),
                assignment_counts AS (
                    SELECT wp_users.ID AS user_id, 
                        COUNT(*) AS assignment_count 
                    FROM wp_comments 
                    JOIN wp_users ON wp_comments.user_ID = wp_users.ID 
                    WHERE wp_comments.comment_type = "tutor_assignment"
                    AND wp_comments.comment_post_ID IN (8260, 8272, 8271) 
                    GROUP BY wp_users.ID
                ),
                lesson_completions AS (
                    SELECT user_id, 
                        COUNT(*) AS lesson_complete 
                    FROM wp_usermeta 
                    WHERE meta_key IN ( 
                        "_tutor_completed_lesson_id_8045", 
                        "_tutor_completed_lesson_id_8046", 
                        "_tutor_completed_lesson_id_8047", 
                        "_tutor_completed_lesson_id_8048", 
                        "_tutor_completed_lesson_id_8049", 
                        "_tutor_completed_lesson_id_8050", 
                        "_tutor_completed_lesson_id_8051", 
                        "_tutor_completed_lesson_id_8052", 
                        "_tutor_completed_lesson_id_8053", 
                        "_tutor_completed_lesson_id_8054" 
                    ) 
                    GROUP BY user_id
                )
                
                SELECT 
                    q.user_id, 
                    q.total_quizzes, 
                    a.assignment_count, 
                    l.lesson_complete,
                    COALESCE(q.total_quizzes, 0) + COALESCE(a.assignment_count, 0) + COALESCE(l.lesson_complete, 0) AS total,
                    concat(round((COALESCE(q.total_quizzes, 0) + COALESCE(a.assignment_count, 0) + COALESCE(l.lesson_complete, 0)) / 24.0 * 100 , 0), "%") AS percentage
                FROM 
                    quiz_counts q
                LEFT JOIN 
                    assignment_counts a ON q.user_id = a.user_id
                LEFT JOIN 
                    lesson_completions l ON q.user_id = l.user_id 
            ) as `progress` ON _q_.`base.User_ID` = `progress`.`user_ID`
        LEFT JOIN 
            (
                SELECT 
                    quiz_id, 
                    user_id, 
                    mark2,
                    CASE
                        WHEN mark2 IS NULL THEN "pending"
                        WHEN mark2 <= 50 THEN "failed"
                        ELSE "success"
                    END as status,
                    CASE
                        WHEN mark2 IS NULL THEN "pending"
                        ELSE "completed"
                    END as completed
                FROM (
                    SELECT 
                        quiz_id, 
                        wp_tutor_quiz_attempt_answers.user_id, 
                        round(sum(achieved_mark) / sum(question_mark) * 100, 2) as mark2
                    FROM 
                        wp_users
                    JOIN 
                        wp_tutor_quiz_attempt_answers 
                    ON 
                        wp_users.id = wp_tutor_quiz_attempt_answers.user_id
                    WHERE 
                        quiz_id = "9448"
                    GROUP BY 
                        quiz_id, 
                        wp_tutor_quiz_attempt_answers.user_id
                ) as subquery
            ) as `complete_status` ON _q_.`base.User_ID` = `complete_status`.`user_ID`
        LEFT JOIN 
            (
                SELECT 
                    a.quiz_id, 
                    a.user_id, 
                    a.mark2,
                    CASE
                        WHEN a.mark2 <= "50%" THEN "failed"
                        ELSE "success"
                    END AS status,
                    CASE
                        WHEN a.mark2 IS NULL THEN NULL
                        ELSE b.completion_date
                    END AS completion_date
                FROM (
                    SELECT 
                        quiz_id, 
                        wp_tutor_quiz_attempt_answers.user_id, 
                        CONCAT(ROUND(SUM(achieved_mark) / SUM(question_mark) * 100, 2), "%") AS mark2
                    FROM 
                        wp_users
                    JOIN 
                        wp_tutor_quiz_attempt_answers 
                    ON 
                        wp_users.id = wp_tutor_quiz_attempt_answers.user_id
                    WHERE 
                        quiz_id = "9448"
                    GROUP BY 
                        quiz_id, 
                        wp_tutor_quiz_attempt_answers.user_id
                ) AS a
                LEFT JOIN (
                    SELECT 
                        user_id, 
                        DATE(attempt_ended_at) AS completion_date,
                        quiz_id
                    FROM 
                        wp_tutor_quiz_attempts 
                    WHERE 
                        quiz_id = "9448"
                ) AS b
                ON 
                    a.user_id = b.user_id AND a.quiz_id = b.quiz_id
            ) as `comp` ON _q_.`base.User_ID` = `comp`.`user_ID`
        LEFT JOIN 
            (
                SELECT 
                    user_id, 
                    comment_post_ID, 
                    COUNT(*) AS comment_count
                FROM 
                    wp_comments
                WHERE 
                    comment_post_ID = "7619"
                    AND comment_type = "course_completed"
                GROUP BY 
                    user_id, 
                    comment_post_ID
                HAVING 
                    COUNT(*) >= 1
            ) as `retry` ON _q_.`base.User_ID` = `retry`.`user_ID`';

    // Get results from the database
    $results = $wpdb->get_results( $query );

    // Create a new Spreadsheet object
    $spreadsheet = new Spreadsheet();

    // Get the active sheet
    $sheet = $spreadsheet->getActiveSheet();

    // Set headers
    $headers = array_keys( (array) $results[0] );
    $col = 1;
    foreach ( $headers as $header ) {
        $sheet->setCellValueByColumnAndRow( $col, 1, $header );
        $col++;
    }

    // Add data to the Excel file
    $row = 2;
    foreach ( $results as $result ) {
        $col = 1;
        foreach ( $result as $value ) {
            $sheet->setCellValueByColumnAndRow( $col, $row, $value );
            $col++;
        }
        $row++;
    }

    // Set headers for download
    header( 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
    header( 'Content-Disposition: attachment;filename="courses.xlsx"' );
    header( 'Cache-Control: max-age=0' );

    // Output the Excel file to browser
    $writer = new Xlsx( $spreadsheet );
    $writer->save( 'php://output' );

    exit;
}

// Function to render export page content
function render_export_page() {
    ?>
    <div class="wrap">
        <h1>Export Courses to Excel</h1>
        <form method="post">
            <?php wp_nonce_field( 'export_courses_nonce', 'export_courses_nonce' ); ?>
            <input type="hidden" name="export_courses" value="1" />
            <?php submit_button( 'Export Courses' ); ?>
        </form>
    </div>
    <?php
}

// Hook to handle export process when form is submitted
function handle_export_process() {
    if ( isset( $_POST['export_courses'] ) && isset( $_POST['export_courses_nonce'] ) && wp_verify_nonce( $_POST['export_courses_nonce'], 'export_courses_nonce' ) ) {
        export_courses_to_excel();
    }
}

// Hook to add admin page
function add_export_page() {
    add_menu_page(
        'Export report',
        'Export report',
        'manage_options',
        'export-report',
        'render_export_page'
    );
}
add_action( 'admin_menu', 'add_export_page' );

// Hook to handle export process
add_action( 'admin_init', 'handle_export_process' );
?>