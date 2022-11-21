function send_notification_email(e) 
{    
  var sheet = e.source.getActiveSheet();


  // THIS PART IS FOR ALL ORDERS TAB - IF A ROW HAS ATTENTION ON IT, BELOW CODE EXECUTES. 
  if(sheet.getName() == "All orders")
  { 
    // looks up for columns
    const status_cell = e.source.getSheetByName("All orders").getRange(e.range.rowStart,11,1,1).getValues()
    // gets row of event
    var row = e.range.getRow();
    // if event is true
    if(status_cell[0][0] == true)
    {
      // fills cell with date when checked box
      e.source.getActiveSheet().getRange(row,12).setValue(new Date());
    }
    // looks up for columns
    const sh = e.source.getSheetByName("All orders").getRange(e.range.rowStart,15,1,1).getValues()
    //if column matches with Attention, runs 
    if(sh[0] == "ATTENTION!")
    {
      const shData = e.source.getSheetByName("All orders").getRange(e.range.rowStart,1,1,14).getValues();
      let machine_Code = shData[0][0];
      let scheduled_Date = shData[0][1];
      let start_Time = shData[0][2];
      let first_Order	= shData[0][3];
      let product_Code	= shData[0][4];
      let wO_Number = shData[0][5];
      let order_Qty = shData[0][6];
      let passes	= shData[0][7];
      let required_Date = shData[0][8];
      let vTR_No = shData[0][9];
      let status = shData[0][10];
      let time_order_Finished = shData[0][11];
      let good_Passes_Done = shData[0][12];
      let percentage_done = shData[0][13];

      let msg =  " Machine code: " + machine_Code +"\n"+ " Schedule Date: " + scheduled_Date +"\n"+ " Start time: " + start_Time +"\n"+ " First order: " + first_Order +"\n"+ " Product Code: " + product_Code +"\n"+ " Order Qty: " + order_Qty +"\n"+ " Wo Number " + wO_Number+"\n"+ " Passes: " + passes +"\n"+ " Required Date: " + required_Date +"\n"+ " VTR No: " + vTR_No +"\n"+ " Status: " + status +"\n"+ " Time Order Finished: " + time_order_Finished +"\n"+ " Good Passes Done: " + good_Passes_Done +"\n"+ " Percertage done: " + percentage_done;
      Logger.log(msg);
      //GmailApp.sendEmail("joteraw561@sopulit.com", "The following order requires attention: " , msg);
    }
  }
  
  // THIS PART IS FOR  ATTENTION ORDERS TAB - NOTIFIES FEEDBACK OF ABOVE
  
  else if(sheet.getName() == "Attention orders")
  { 
    
    const sh_feedback = e.source.getActiveSheet().getRange(e.range.rowStart,19,1,1).getValues()
    if(sh_feedback[0] == "COMPLETE" || sh_feedback[0] == "IN PROGRESS, PLEASE WAIT")
    {
      const sh_feedbackData = e.source.getSheetByName("Attention orders").getRange(e.range.rowStart,1,1,20).getValues();
      let machine_Code = sh_feedbackData[0][0];
      let scheduled_Date = sh_feedbackData[0][1];
      let start_Time = sh_feedbackData[0][2];
      let first_Order	= sh_feedbackData[0][3];
      let product_Code	= sh_feedbackData[0][4];
      let wO_Number = sh_feedbackData[0][5];
      let order_Qty = sh_feedbackData[0][6];
      let passes	= sh_feedbackData[0][7];
      let required_Date = sh_feedbackData[0][8];
      let vTR_No = sh_feedbackData[0][9];
      let status = sh_feedbackData[0][10];
      let time_order_Finished = sh_feedbackData[0][11];
      let good_Passes_Done = sh_feedbackData[0][12];
      let percentage_done = sh_feedbackData[0][13];
      let issue_from_Production = sh_feedbackData[0][15];
      let abaca_completed = sh_feedbackData[0][16];
      let planning_dept_feedback = sh_feedbackData[0][16];
      let planning_dept_status = sh_feedbackData[0][17];

      let planning_feedback =  " For order: " + "\n\n" + " Machine code: " + machine_Code +"\n"+ " Schedule Date: " + scheduled_Date +"\n"+ " Start time: " + start_Time +"\n"+ " First order: " + first_Order +"\n"+ " Product Code: " + product_Code +"\n"+ " Order Qty: " + order_Qty +"\n"+ " Wo Number " + wO_Number+"\n"+ " Passes: " + passes +"\n"+ " Required Date: " + required_Date +"\n"+ " VTR No: " + vTR_No +"\n"+ " Status: " + status +"\n"+ " Time Order Finished: " + time_order_Finished +"\n"+ " Good Passes Done: " + good_Passes_Done +"\n"+ " Percertage done: " + percentage_done + "\n\n" + " Feedback from Planning department below: " + "\n\n" + "Issue from production: " + issue_from_Production + "\n" + "Order completed on Abaca: " + abaca_completed + "\n" + "Planning Dept. Feedback: " + planning_dept_feedback + "\n" + "Status: " + planning_dept_status;
      Logger.log(planning_feedback);
      //GmailApp.sendEmail("najil62448@sunetoa.com", "Feedback from planning department: " , planning_feedback);
    }
  }

  // THIS BLOCK IS FOR WAREHOUSE TAB

  else if(sheet.getName() == "Warehouse")
  { 
    // looks up for column
    const sh = e.source.getActiveSheet().getRange(e.range.rowStart,17,1,1).getValues()
    //if column matches, runs 
    if(sh[0][0] == true)
    {
      const shData = e.source.getSheetByName("Warehouse").getRange(e.range.rowStart,1,1,28).getValues();
      let abaca_Time = shData[0][0];
      let depart_Date = shData[0][1];
      let driver_Name = shData[0][2];
      let vehicle_Trip_Number	= shData[0][3];
      let vehicle_Description	= shData[0][4];
      let drop_Date = shData[0][5];
      let drop_Time = shData[0][6];
      let book_In_Note1	= shData[0][7];
      let book_In_Note2 = shData[0][8];
      let delivery_Name = shData[0][9];
      let postcode = shData[0][10];
      let pallets = shData[0][11];
      let product_Code = shData[0][12]
      let sales_Order_No = shData[0][13];
      let works_Order_No = shData[0][14];
      let loaded_with_Balance = shData[0][16];
      let works_Order_Stages = shData[0][18];
      let wO_Due_Date = shData[0][19];
      let wO_Due_Time = shData[0][20];
      let date_Status = shData[0][21];
      let required_Qty = shData[0][22];
      let trip_Status = shData[0][23];
      let wO_Made_Qty = shData[0][24];
      let customer_Name = shData[0][25];
      let trip_Comment_1 = shData[0][26];
      let trip_Comment_2 = shData[0][27];

      let msg =  
      " Abaca Time: " + abaca_Time +"\n"+ 
      " Depart_Date: " + depart_Date +"\n"+ 
      " Driver_Name: " + driver_Name +"\n"+ 
      " Vehicle_Trip_Number: " + vehicle_Trip_Number +"\n"+ 
      " Vehicle_Description: " + vehicle_Description +"\n"+ 
      " Drop_Date: " + drop_Date +"\n"+ 
      " Drop_Time " + drop_Time+"\n"+ 
      " Book_In_Note1: " + book_In_Note1 +"\n"+ 
      " Book_In_Note2: " + book_In_Note2 +"\n"+ 
      " Delivery_Name: " + delivery_Name +"\n"+ 
      " Postcode: " + postcode +"\n"+ 
      " Pallets: " + pallets +"\n"+ 
      " Product_Code: " +  product_Code+"\n"+ 
      " Sales_Order_No: " + sales_Order_No +"\n"+ 
      " Works_Order_No: " + works_Order_No +"\n"+ 
      " Sales_Order_No: " + sales_Order_No +"\n"+ 
      " Loaded_with_Balance: " + loaded_with_Balance +"\n"+ 
      " Works_Order_Stages: " + works_Order_Stages +"\n"+
      " WO_Due_Date: " + wO_Due_Date +"\n"+ 
      " WO_Due_Time: " + wO_Due_Time +"\n"+ 
      " Date_Status: " + date_Status +"\n"+ 
      " Required_Qty: " + required_Qty +"\n"+ 
      " Trip_Status: " + trip_Status +"\n"+ 
      " WO_Made_Qty: " + wO_Made_Qty +"\n"+ 
      " Customer_Name: " + customer_Name +"\n"+ 
      " Trip_Comment_1: " + trip_Comment_1 +"\n"+ 
      " Trip_Comment_2: " + trip_Comment_2 +"\n";
      Logger.log(msg);
      //GmailApp.sendEmail("najil62448@sunetoa.com", "The following order was loaded with balance: " , msg);
    }
  }
  // FEEDBACK FROM WAREHOUSE STATING DETAILS OF PENDING ORDER. EVENT HAPPENS WHEN LAST CELL (DATE) IS FILLED - NEXT BLOCK IS FOR CONFIRMATION, WHEN NEXT CELL IS TICKED.
  else if(sheet.getName() == "Warehouse with balance")
  { 
    // looks up for column
    const sh = e.source.getActiveSheet().getRange(e.range.rowStart,23,1,1).getValues()
    //if column matches, runs 
    if(sh[0][0] != 0)
    {
      const shData = e.source.getSheetByName("Warehouse with balance").getRange(e.range.rowStart,1,1,23).getValues();
      let abaca_Time = shData[0][0];
      let depart_Date = shData[0][1];
      let driver_Name = shData[0][2];
      let vehicle_Trip_Number	= shData[0][3];
      let vehicle_Description	= shData[0][4];
      let drop_Date = shData[0][5];
      let drop_Time = shData[0][6];
      let book_In_Note1	= shData[0][7];
      let book_In_Note2 = shData[0][8];
      let delivery_Name = shData[0][9];
      let postcode = shData[0][10];
      let pallets = shData[0][11];
      let product_Code = shData[0][12]
      let sales_Order_No = shData[0][13];
      let works_Order_No = shData[0][14];
      let loaded_with_Balance = shData[0][16];
      let no_of_Pallets = shData[0][18];
      let location = shData[0][19];
      let re_scheduled_delivery_date = shData[0][20];
      let nEW_VTR = shData[0][21];
      let reScheduled_date_Changed_On_Abaca = shData[0][22];
      
      let msg =  
      " Abaca Time: " + abaca_Time +"\n"+ 
      " Depart_Date: " + depart_Date +"\n"+ 
      " Driver_Name: " + driver_Name +"\n"+ 
      " Vehicle_Trip_Number: " + vehicle_Trip_Number +"\n"+ 
      " Vehicle_Description: " + vehicle_Description +"\n"+ 
      " Drop_Date: " + drop_Date +"\n"+ 
      " Drop_Time " + drop_Time+"\n"+ 
      " Book_In_Note1: " + book_In_Note1 +"\n"+ 
      " Book_In_Note2: " + book_In_Note2 +"\n"+ 
      " Delivery_Name: " + delivery_Name +"\n"+ 
      " Postcode: " + postcode +"\n"+ 
      " Pallets: " + pallets +"\n"+ 
      " Product_Code: " +  product_Code+"\n"+ 
      " Sales_Order_No: " + sales_Order_No +"\n"+ 
      " Works_Order_No: " + works_Order_No +"\n"+ 
      " Sales_Order_No: " + sales_Order_No +"\n"+ 
      " Loaded_with_Balance: " + loaded_with_Balance +"\n\n"+
      "FEEDBACK FROM WAREHOUSE" + "\n\n"+
      " No of Pallets: " + no_of_Pallets +"\n"+
      " Location: " + location +"\n"+ 
      " Re scheduled delivery date: " + re_scheduled_delivery_date +"\n"+ 
      " NEW VTR: " + nEW_VTR +"\n"+ 
      " ReScheduled date Changed On Abaca: " + reScheduled_date_Changed_On_Abaca;
      Logger.log(msg);
      //GmailApp.sendEmail("najil62448@sunetoa.com", "Feedback from warehouse: " , msg);
    }
  }
  // FEEDBACK FROM WAREHOUSE FOR ORDER COMPLETION
  else if(sheet.getName() == "Warehouse with balance")
  { 
    // looks up for column
    const sh = e.source.getActiveSheet().getRange(e.range.rowStart,24,1,1).getValues()
    //if column matches, runs 
    if(sh[0][0] == true)
    {
      const shData = e.source.getSheetByName("Warehouse with balance").getRange(e.range.rowStart,1,1,23).getValues();
      let abaca_Time = shData[0][0];
      let depart_Date = shData[0][1];
      let driver_Name = shData[0][2];
      let vehicle_Trip_Number	= shData[0][3];
      let vehicle_Description	= shData[0][4];
      let drop_Date = shData[0][5];
      let drop_Time = shData[0][6];
      let book_In_Note1	= shData[0][7];
      let book_In_Note2 = shData[0][8];
      let delivery_Name = shData[0][9];
      let postcode = shData[0][10];
      let pallets = shData[0][11];
      let product_Code = shData[0][12]
      let sales_Order_No = shData[0][13];
      let works_Order_No = shData[0][14];
      let loaded_with_Balance = shData[0][16];
      let no_of_Pallets = shData[0][18];
      let location = shData[0][19];
      let re_scheduled_delivery_date = shData[0][20];
      let nEW_VTR = shData[0][21];
      let reScheduled_date_Changed_On_Abaca = shData[0][22];
      
      let msg =  
      " Abaca Time: " + abaca_Time +"\n"+ 
      " Depart_Date: " + depart_Date +"\n"+ 
      " Driver_Name: " + driver_Name +"\n"+ 
      " Vehicle_Trip_Number: " + vehicle_Trip_Number +"\n"+ 
      " Vehicle_Description: " + vehicle_Description +"\n"+ 
      " Drop_Date: " + drop_Date +"\n"+ 
      " Drop_Time " + drop_Time+"\n"+ 
      " Book_In_Note1: " + book_In_Note1 +"\n"+ 
      " Book_In_Note2: " + book_In_Note2 +"\n"+ 
      " Delivery_Name: " + delivery_Name +"\n"+ 
      " Postcode: " + postcode +"\n"+ 
      " Pallets: " + pallets +"\n"+ 
      " Product_Code: " +  product_Code+"\n"+ 
      " Sales_Order_No: " + sales_Order_No +"\n"+ 
      " Works_Order_No: " + works_Order_No +"\n"+ 
      " Sales_Order_No: " + sales_Order_No +"\n"+ 
      " Loaded_with_Balance: " + loaded_with_Balance +"\n\n"+
      "FEEDBACK FROM WAREHOUSE" + "\n\n"+
      " No of Pallets: " + no_of_Pallets +"\n"+
      " Location: " + location +"\n"+ 
      " Re scheduled delivery date: " + re_scheduled_delivery_date +"\n"+ 
      " NEW VTR: " + nEW_VTR +"\n"+ 
      " ReScheduled date Changed On Abaca: " + reScheduled_date_Changed_On_Abaca;
      Logger.log(msg);
      //GmailApp.sendEmail("najil62448@sunetoa.com", "Feedback from warehouse: Below order was completed: " , msg);
    }
  }  
  
}
