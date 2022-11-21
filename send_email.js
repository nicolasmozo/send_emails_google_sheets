function send_notification_email(e) 
{    
  var sheet = e.source.getActiveSheet();

  // THIS PART IS FOR ALL ORDERS TAB - IF A ROW HAS ATTENTION ON IT, BELOW CODE EXECUTES. 
  if(sheet.getName() == "All orders")
  {
    // looks up for column
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
    //Logger.log(sh_feedback);
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
      GmailApp.sendEmail("najil62448@sunetoa.com", "Feedback from planning department: " , planning_feedback);
    }
  }
}