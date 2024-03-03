function offHourEmailResponse() {

  // google sheet information set up;

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = spreadSheet.getSheetByName("SETUP").getRange("A3:F3").getValues();

  const activationStatus = setupSheet[0][0];
  const workDayStartHour = setupSheet[0][1];
  const workDayEndHour = setupSheet[0][2];
  const customizeMessage = setupSheet[0][4];

  const myEmail = Session.getActiveUser().getEmail();

  // google sheet record sheet:

  if(workDayStartHour.length === 0 || workDayEndHour.length === 0){
    console.log("info not valid");
    return
  }

  // const weekendFrom = setupSheet[0][3];
  // const weekendTo = setupSheet[0][4];

  const whiteListRaw = setupSheet[0][3].toString().split(",");
  const whiteListSorted = whiteListRaw.map((email)=> email.trim());           // get ride of white space if there's any

  const dateInfo = new Date()
  const today = dateInfo.getDay();
  const timeNow = dateInfo.getTime() / (60*1000);                             //check times in minutes
  const dwightSpecialMessage = "In order to maintain a healthy work-life balance and ensure the highest quality of service during operational times, this is adhered to. ";

  var inboxThreads = GmailApp.getInboxThreads(0,5);                          // get Email information
  
  if(activationStatus === "HOLIDAYS") {
      
      for(let i =0;i < inboxThreads.length; i ++){

        const lastMessageIndex = inboxThreads[i].getMessages().length -1;                       // get last message Index
        let messageSender = inboxThreads[i].getMessages()[lastMessageIndex].getFrom();          // get who sent the email

        let senderEmail;                    
        if(messageSender.includes("<")){                              // could be 2 variation of email addresses, to avoid errors.
            senderEmail = messageSender.split("<")[1].slice(0,-1);
        } else {
            senderEmail = messageSender;
        }

        if(senderEmail === myEmail){
          console.log("The email is from me, do nothing")
          continue
        }

        let pastEmailCount = scanAndCount(senderEmail);

        const inboxEmailTime = inboxThreads[i].getMessages()[lastMessageIndex].getDate().getTime() / (60*1000);
        const lessThan60min = timeNow - inboxEmailTime < 60;            // from current time to email time less than 60
        const emailSubject = inboxThreads[i].getMessages()[lastMessageIndex].getSubject();

        // let emailCountReminder = "";
        // if(pastEmailCount >0){
        //   emailCountReminder = "\nIt appears that you have sent email to ${myEmail} \n\n${pastEmailCount} times \n\n this school year during non-working hours.";
        // }

        if(!whiteListSorted.includes(senderEmail)){                     // not on the exempt list, start processing conditions
            inboxThreads[i].markRead();

            if(lessThan60min) {
                let emailBody;
                if(customizeMessage.length !== 0){
                  emailBody = customizeMessage;
                }else{
                  emailBody = `Greetings

                      \nThe recipient of this email ${myEmail} is currently unavailable during holidays. ${senderEmail.includes("@yourOrgEmail.com")?dwightSpecialMessage:""}
                      \nWork Hours: 
                      \nMonday to Friday: 8:00 AM to 5:00 PM 
                      \nExcludes: Public Holidays and Organizational Holidays.
                      \nPlease note that any emails received outside of these hours will be marked as read automatically and will not trigger notifications. Your understandings are greatly appreciated.
                      \nThank you
                      \nWarmest Regards`;
                  }
                // SEND EMAIL
                inboxThreads[i].getMessages()[lastMessageIndex].reply(emailBody);
                // SENT EMAIL
                console.log("Holiday Email has been sent to: "+senderEmail)

                const emailTime = inboxThreads[i].getMessages()[lastMessageIndex].getDate().toLocaleString();
                // write to spreadsheet
                writeToRecord(senderEmail,myEmail,emailTime,emailSubject);

            }else{
              console.log("Email possibily has been sent to "+senderEmail +"or it's more than 60 mins old");
            }
        }else{
            writeToRecord(senderEmail,myEmail,emailTime,emailSubject);
            continue
        }
      }

  } else if (activationStatus === "ACTIVE") {
      console.log(dateInfo.getDay())
      if(dateInfo.getDay() !== 6 && dateInfo.getDay() !== 0){
          if(dateInfo.getHours() < workDayEndHour && dateInfo.getHours() > workDayStartHour){
            return
          }
      }

    // starting to loop through the emails - 
        for(let i = 0; i < inboxThreads.length; i ++){
          
            const lastMessageIndex = inboxThreads[i].getMessages().length -1;         // get last message index
            let messageSender = inboxThreads[i].getMessages()[lastMessageIndex].getFrom();     // get email sender information
            let senderEmail;                                            // clean up the email address
            if(messageSender.includes("<")){                            // check variation of email format
                senderEmail = messageSender.split("<")[1].slice(0, -1);
            }else{
                senderEmail = messageSender;
            }
            // check if it's repetitive
            let pastEmailCount = scanAndCount(senderEmail);

            // let emailCountReminder = "";
            // if(pastEmailCount >0){
            //   emailCountReminder = "\nIt appears that you have sent email to ${myEmail} \n\n${pastEmailCount} times \n\n this school year during non-working hours.";
            // }
            const inboxEmailTime = inboxThreads[i].getMessages()[lastMessageIndex].getDate().getTime() / (60*1000);  
            const emailSubject = inboxThreads[i].getMessages()[lastMessageIndex].getSubject();
            const emailTime = inboxThreads[i].getMessages()[lastMessageIndex].getDate().toLocaleString();

            // get time of email received
            // setting up conditions       
            const lessThan60min = timeNow - inboxEmailTime < 60;
            const laterThanEndTime = inboxThreads[i].getMessages()[lastMessageIndex].getDate().getHours() >= workDayEndHour;
            const earlierThanStartTime = inboxThreads[i].getMessages()[lastMessageIndex].getDate().getHours() < workDayStartHour;

            if(senderEmail === myEmail){
              // writeToRecord(senderEmail,myEmail,emailTime,emailSubject);
              console.log("The email is form me, do nothing")
              continue
            }

            if(!whiteListSorted.includes(senderEmail)){                     // if not on the exempt list, start processing conditions
                
              inboxThreads[i].markRead();                          // mark the email as read automatically;
                // Weekend CASE
              if(today === 6 || today ===0) {                           // if it's weekend , proceed to weekend email
                  console.log("weekend email response");                
                let emailBody;
                if(customizeMessage.length !== 0){
                    emailBody = customizeMessage;
                }else{
                    emailBody = `Greetings
                        \nThe recipient of this email ${myEmail} is currently unavailable during weekends.${senderEmail.includes("@yourOrgEmail.com")?dwightSpecialMessage:""}
                        \nWork Hours: 
                        \nMonday to Friday: 8:00 AM to 5:00 PM 
                        \nExcludes: Public Holidays and Organizational Holidays.
                        \nPlease note that any emails received outside of these hours will be marked as read automatically and will not trigger notifications. Your understandings are greatly appreciated.
                        \nThank you.
                        \nWarmest Regards`;
                }
                  // SEND EMAIL
                if(lessThan60min){
                    // respond email
                    inboxThreads[i].getMessages()[lastMessageIndex].reply(emailBody);
                    // record to spreadsheet;       
                  const emailTime = inboxThreads[i].getMessages()[lastMessageIndex].getDate().toLocaleString();
                  writeToRecord(senderEmail,myEmail,emailTime,emailSubject);

                }else{
                  console.log("Do nothing");
                }
                  // SEND EMAIL
              } else {
                console.log("weekday email response:");
                console.log(lessThan60min, laterThanEndTime,earlierThanStartTime)
                console.log(senderEmail)
                                      // 17           6              7
                if(lessThan60min && (laterThanEndTime || earlierThanStartTime)) {   
                                             // check hours for work hours
                    console.log("Boom, outside work hours, email launched to" + senderEmail);
                    //check record
                    const emailTime = inboxThreads[i].getMessages()[lastMessageIndex].getDate().toLocaleString();
                    // record to spreadsheet;       
                    writeToRecord(senderEmail,myEmail,emailTime,emailSubject);

                    let emailBody;
                    if(customizeMessage.length !== 0){
                      emailBody = customizeMessage;
                    }else{
                      
                      emailBody = `Greetings
                        \nThe recipient of this email ${myEmail} is currently unavailable outside of designated working hours. ${senderEmail.includes("@yourOrgEmail.com")?dwightSpecialMessage:""}
                        \nWork Hours: 
                        \n\Monday to Friday: 8:00 AM to 5:00 PM 
                        \nExcludes: Public Holidays and Organizational Holidays.
                        \nPlease note that any emails received outside of these hours will be marked as read automatically and will not trigger notifications. Your understandings are greatly appreciated.
                        \nThank you
                        \nWarmest Regards`;
                    }
                    // SEND EMAIL
                        inboxThreads[i].getMessages()[lastMessageIndex].reply(emailBody);
                      // SEND EMAIL   
                  }else{
                    console.log("Do nothing");
                  }
              }                                         
            }else{
                  // writeToRecord(senderEmail,myEmail,emailTime,emailSubject);
                  console.log(senderEmail + " is on the whitelist")
                  continue
            }
        }
  } else if (activationStatus === "INACTIVE") {
      console.log(activationStatus)
      console.log("It's inactive");
      return
  }
}

function writeToRecord(from,to,time,subject){
    let spreadsheet = SpreadsheetApp.openById("1Q5paDvQ6pug-5prvo2AC4uIRAy-y5sdfXpvft0FdJx4")
    let recordSheet = spreadsheet.getSheetByName("Records");
    if(to.includes("@dwight.or.kr")){
      recordSheet.appendRow([from,to,time,subject]);
    }
}

function scanAndCount(email){
  let spreadsheet = SpreadsheetApp.openById("1Q5paDvQ6pug-5prvo2AC4uIRAy-y5sdfXpvft0FdJx4");
  let recordSheet = spreadsheet.getSheetByName("Records").getDataRange().getValues();

  let count = 0;
  for(let i = 0; i < recordSheet.length; i ++){
    if(recordSheet[i][0] === email){
      count ++
    }
  }
  return count;
}








