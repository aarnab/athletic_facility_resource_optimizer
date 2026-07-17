/* 
   PART 1: Generating mock data
*/

DATA WORK.Gym_Inventory;
    /* Define columns for equipment and usage stats */
    LENGTH Item_Name $ 30 Training_Focus $ 15;
    INPUT Equip_ID Item_Name $ Daily_Usage_Hrs Days_Since_Service Training_Focus $;
    DATALINES;
101 Powerlifting_Barbell 8 120 Powerlifting
102 Flat_Bench 6 150 Hypertrophy
103 Tug_of_War_Rope 2 45 Powersports
104 Cable_Crossover 10 200 Hypertrophy
105 Squat_Rack 9 90 Powerlifting
;
RUN;

/* 
   PART 2: Running the analytical model
*/

DATA WORK.Maintenance_Schedule;
    SET WORK.Gym_Inventory;
    
    /* Calculate a 'Wear and Tear Score' based on daily hours and days since last check */
    Wear_Score = Daily_Usage_Hrs * (Days_Since_Service / 30);
    
    /* Heavy powerlifting equipment requires faster servicing */
    IF Training_Focus = 'Powerlifting' THEN Wear_Score = Wear_Score * 1.2;
    
    /* Determine Status */
    IF Wear_Score > 40 THEN Status = 'CRITICAL: Service Immediately';
    ELSE IF Wear_Score > 20 THEN Status = 'Warning: Schedule Service';
    ELSE Status = 'Optimal Condition';
RUN;

/* Summarising the maintenance priorities */
PROC SQL;
    CREATE TABLE WORK.Action_Required AS
    SELECT Item_Name, Training_Focus, Wear_Score, Status
    FROM WORK.Maintenance_Schedule
    ORDER BY Wear_Score DESC;
QUIT;

/* 
   PART 3: Displaying the results
*/
PROC PRINT DATA=WORK.Action_Required;
    TITLE "Facility Operations: Equipment Maintenance Forecast";
RUN;
