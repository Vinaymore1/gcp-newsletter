/**
 * Creates the newsletter data spreadsheet structure
 */
function createNewsletterSpreadsheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create master sheet for basic info
    createBasicInfoSheet(ss);
    
    // Create sheets for each section
    createSectionSheet(ss, "LeadershipMessage", ["title", "content"], false);
    createSectionSheet(ss, "FeaturedUpdates", ["title", "description", "details"], true);
    createSectionSheet(ss, "Initiatives", ["title", "description", "image", "details"], true);
    createSectionSheet(ss, "SecurityUpdate", ["title", "description", "details", "KeyChanges"], false);
    createSectionSheet(ss, "TechnicalTip", ["content"], false);
    createSectionSheet(ss, "Learnings", ["title", "description", "link","buttonText"], true);
    createSectionSheet(ss, "Events", ["title", "description", "link","buttonText"], true);
    createSectionSheet(ss, "Certifications", ["name", "achievement", "date", "image"], true);
    createSectionSheet(ss, "Shoutouts", ["name", "contribution", "image"], true);
    createSectionSheet(ss, "NextMonthFocus", ["title", "description", "icon", "iconClass"], true);
    createSectionSheet(ss, "SocialLinks", ["url", "image"], true);
    
    // Create a sheet for adding a feedback message
    createFeedbackSheet(ss);
    
    // Create JSON Generation sheet
    createJsonGeneratorSheet(ss);
    
    // Set the first sheet as active
    ss.setActiveSheet(ss.getSheets()[0]);
    
    // After creating the structure, populate with sample data
    populateWithSampleData(ss);
  }
  
  /**
   * Creates the basic info sheet
   */
  function createBasicInfoSheet(ss) {
    const sheet = ss.getSheetByName("BasicInfo") || ss.insertSheet("BasicInfo");
    sheet.clear();
    
    // Set up headers and initial values
    sheet.getRange("A1:B1").setValues([["Field", "Value"]]);
    sheet.getRange("A2:A4").setValues([["Issue Number"], ["Issue date"], ["Current Year"]]);
    sheet.getRange("B2:B4").setValues([[1], ["April 2025"], ["2025"]]);
    
    // Format the sheet
    sheet.getRange("A1:B1").setFontWeight("bold");
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 300);
    sheet.getRange("A1:B4").setBorder(true, true, true, true, true, true);
  }
  
  /**
   * Creates a sheet for a specific section
   */
  function createSectionSheet(ss, sectionName, fields, isArray) {
    const sheet = ss.getSheetByName(sectionName) || ss.insertSheet(sectionName);
    sheet.clear();
    
    // Set up headers
    const headers = ["Entry ID", ...fields, "Active"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Add example data for first row
    const exampleRow = [1, ...Array(fields.length).fill("Example"), true];
    sheet.getRange(2, 1, 1, exampleRow.length).setValues([exampleRow]);
    
    // Format the sheet
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 80); // Entry ID
    for (let i = 0; i < fields.length; i++) {
      sheet.setColumnWidth(i + 2, fields[i] === "details" ? 400 : 200);
    }
    sheet.setColumnWidth(headers.length, 80); // Active column
    
    // Add a note to the sheet
    if (isArray) {
      sheet.getRange("A1").setNote("This section supports multiple entries. Add a new row for each entry and set 'Active' to TRUE to include it in the JSON.");
    } else {
      sheet.getRange("A1").setNote("This section supports only one entry. Only the first row with 'Active' set to TRUE will be used.");
    }
  }
  
  /**
   * Creates the feedback text sheet
   */
  function createFeedbackSheet(ss) {
    const sheet = ss.getSheetByName("Feedback") || ss.insertSheet("Feedback");
    sheet.clear();
    
    // Set up headers and initial values
    sheet.getRange("A1").setValue("Feedback Text");
    sheet.getRange("A2").setValue("Your insights matter to us! Share your thoughts on our newsletter content, design, and delivery frequency. We're constantly working to improve your experience and provide content that matters most to you. Have ideas for future topics? Let us know what you'd like to see in upcoming editions.");
    
    // Format the sheet
    sheet.getRange("A1").setFontWeight("bold");
    sheet.setColumnWidth(1, 600);
    sheet.setRowHeight(2, 100);
  }
  
  /**
   * Creates a sheet with a button to generate JSON
   */
  function createJsonGeneratorSheet(ss) {
    const sheet = ss.getSheetByName("GenerateJSON") || ss.insertSheet("GenerateJSON");
    sheet.clear();
    
    // Add instructions and button
    sheet.getRange("A1:B1").merge();
    sheet.getRange("A1").setValue("NEWSLETTER JSON GENERATOR").setFontWeight("bold").setHorizontalAlignment("center");
    
    sheet.getRange("A3").setValue("Click the button below to generate JSON:");
    sheet.getRange("A4:B4").merge();
    
    // Create a button using a drawing
    const buttonCell = sheet.getRange("A4");
    buttonCell.setValue("GENERATE JSON")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#4285F4")
      .setFontColor("white");
    
    // Add script assignment instructions
    sheet.getRange("A6:B6").merge();
    sheet.getRange("A6").setValue("To make the button work: Right-click cell A4 → Insert comment → Delete the comment → Click '...' → Edit Assignment → Enter 'generateNewsletterJson'");
    
    // Format output area
    sheet.getRange("A8").setValue("Generated JSON:");
    sheet.getRange("A9:B9").merge();
    sheet.getRange("A9").setWrap(true);
    
    sheet.setColumnWidth(1, 400);
    sheet.setColumnWidth(2, 400);
    sheet.setRowHeight(9, 500); // Make room for JSON output
  }
  
  /**
   * Generates the newsletter JSON from all sheets
   */
  function generateNewsletterJson() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let newsletterData = {};
    
    // Get basic info
    const basicInfoSheet = ss.getSheetByName("BasicInfo");
    const basicInfoValues = basicInfoSheet.getRange("A2:B4").getValues();
    
    newsletterData.issueNumber = basicInfoValues[0][1];
    newsletterData.issuedate = basicInfoValues[1][1];
    newsletterData.currentYear = basicInfoValues[2][1];
    
    // Get leadership message (single entry)
    newsletterData.leadershipMessage = getSheetDataObject(ss, "LeadershipMessage", ["title", "content"]);
    
    // Get featured updates (multiple entries)
    newsletterData.featuredUpdates = getSheetDataArray(ss, "FeaturedUpdates", ["title", "description", "details"]);
    
    // Get initiatives (multiple entries)
    newsletterData.initiatives = getSheetDataArray(ss, "Initiatives", ["title", "description", "image", "details"]);
    
    // Get security update (single entry)
    const securityUpdate = getSheetDataObject(ss, "SecurityUpdate", ["title", "description", "details", "KeyChanges"]);
    if (securityUpdate.KeyChanges) {
      securityUpdate.keyChanges = securityUpdate.KeyChanges.split('\n').map(item => item.trim());
      delete securityUpdate.KeyChanges;
    }
    newsletterData.securityUpdate = securityUpdate;
    
    // Get technical tip (single value)
    const technicalTipSheet = ss.getSheetByName("TechnicalTip");
    const technicalTipData = getSheetDataAsRows(technicalTipSheet, ["content"]);
    if (technicalTipData.length > 0) {
      newsletterData.technicalTip = technicalTipData[0].content;
    }
    
    // Get learnings (multiple entries)
    newsletterData.learnings = getSheetDataArray(ss, "learnings", ["title", "description", "link","buttonText"]);
    
    // Get events (multiple entries)
    newsletterData.events = getSheetDataArray(ss, "events", ["title", "description", "link","buttonText"]);
    
    // Get certifications (multiple entries)
    newsletterData.certifications = getSheetDataArray(ss, "Certifications", ["name", "achievement", "date", "image"]);
    
    // Get shoutouts (multiple entries)
    newsletterData.shoutouts = getSheetDataArray(ss, "Shoutouts", ["name", "contribution", "image"]);
    
    // Get next month focus (multiple entries)
    newsletterData.nextMonthFocus = getSheetDataArray(ss, "NextMonthFocus", ["title", "description", "Icon", "IconClass"]);
    
    // Get feedback text
    const feedbackSheet = ss.getSheetByName("Feedback");
    newsletterData.feedbackText = feedbackSheet.getRange("A2").getValue();
    
    // Get social links (multiple entries)
    newsletterData.socialLinks = getSheetDataArray(ss, "SocialLinks", ["url", "image"]);
    
    // Output the JSON to the generator sheet
    const jsonOutput = JSON.stringify(newsletterData, null, 2);
    const generatorSheet = ss.getSheetByName("GenerateJSON");
    generatorSheet.getRange("A9").setValue(jsonOutput);
    
    // Show a success message
    SpreadsheetApp.getUi().alert("JSON generated successfully!");
    
    // Return the JSON in case we want to use it elsewhere
    return jsonOutput;
  }
  
  /**
   * Helper function to get sheet data as an array of objects with only specified fields
   */
  function getSheetDataArray(ss, sheetName, fieldsToInclude) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const data = getSheetDataAsRows(sheet, fieldsToInclude);
    return data.map(row => {
      const result = {};
      fieldsToInclude.forEach(field => {
        result[field.toLowerCase() === "url" ? "url" : field] = row[field];
      });
      return result;
    });
  }
  
  /**
   * Helper function to get sheet data as a single object with only specified fields
   */
  function getSheetDataObject(ss, sheetName, fieldsToInclude) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {};
    
    const data = getSheetDataAsRows(sheet, fieldsToInclude);
    if (data.length === 0) return {};
    
    // Use only the first active row
    const firstRow = data[0];
    const result = {};
    
    fieldsToInclude.forEach(field => {
      result[field] = firstRow[field];
    });
    
    return result;
  }
  
  /**
   * Helper function to get sheet data as an array of row objects with only specified fields
   */
  function getSheetDataAsRows(sheet, fieldsToInclude) {
    // Add null check to prevent errors
    if (!sheet) {
      console.log("Warning: Sheet not found");
      return []; // Return empty array if sheet is null
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) return []; // Only headers or empty sheet
    
    const headers = values[0];
    const rows = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const active = row[headers.length - 1]; // Last column is 'Active'
      
      // Skip inactive rows
      if (!active) continue;
      
      const rowObj = {};
      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];
        if (header === "Entry ID" || header === "Active") continue;
        if (!fieldsToInclude || fieldsToInclude.includes(header)) {
          rowObj[header] = row[j];
        }
      }
      rows.push(rowObj);
    }
    
    return rows;
  }
  
  /**
   * Populates the sheets with sample data from the JSON
   */
  function populateWithSampleData(ss) {
    // Leadership Message
    const leadershipSheet = ss.getSheetByName("LeadershipMessage");
    leadershipSheet.getRange(2, 1, 1, 4).setValues([[
      1, 
      "Welcome to the first edition of Blue Altair's Google Cloud Platform Community newsletter!", 
      "As proud Google Cloud Partners, we have successfully delivered numerous projects leveraging GCP's cutting-edge technologies. To strengthen this partnership, we emphasize continuous learning, certifications, and hands-on experience with real-world projects. This newsletter is designed to share valuable insights, foster collaboration, expand GCP product knowledge, highlight noteworthy blogs, offer training opportunities, and drive innovation within the GCP ecosystem. This month, we've packed in the latest trends, must-know updates, and a few surprises just for you. Whether you're here for industry news, expert tips, or a dose of inspiration, there's something for everyone!",
      true
    ]]);
    
    // Featured Updates
    const featuredSheet = ss.getSheetByName("FeaturedUpdates");
    featuredSheet.getRange(2, 1, 4, 5).setValues([
      [1, "New Feature Release: Cloud Run for Anthos", "Introducing advanced capabilities for containerized applications.", "BigQuery ML now supports advanced machine learning capabilities, including forecasting models and integration with TensorFlow. These enhancements allow data scientists and analysts to build and deploy predictive models directly within BigQuery, leveraging its scalable infrastructure for faster insights. Businesses can use these capabilities to make data-driven decisions, improve demand forecasting, and enhance AI-driven applications without requiring extensive machine learning expertise. The seamless integration with TensorFlow also enables the use of custom deep learning models, expanding the potential for sophisticated analytics and automation.", true],
      [2, "BigQuery ML Enhancements", "New machine learning capabilities in BigQuery with TensorFlow integration.", "BigQuery ML now supports advanced machine learning capabilities, including forecasting models and integration with TensorFlow. These enhancements allow data scientists and analysts to build and deploy predictive models directly within BigQuery, leveraging its scalable infrastructure for faster insights. Businesses can use these capabilities to make data-driven decisions, improve demand forecasting, and enhance AI-driven applications without requiring extensive machine learning expertise. The seamless integration with TensorFlow also enables the use of custom deep learning models, expanding the potential for sophisticated analytics and automation.", true],
      [3, "Cloud Storage Cost Optimization", "New intelligent tiering options automatically move data.", "The new intelligent tiering options in Cloud Storage automatically analyze access patterns and move data to the most cost-effective storage class without impacting performance. This feature helps businesses optimize storage costs by ensuring frequently accessed data remains in high-performance tiers while infrequently used data is transitioned to lower-cost archival storage. By eliminating the need for manual intervention, organizations can focus on their core operations while benefiting from seamless cost savings and enhanced storage efficiency.", true],
      [4, "Vertex AI Agent Builder", "Build, customize, and deploy AI agents with enterprise-grade security.", "Vertex AI Agent Builder provides enterprise-grade security and governance controls, ensuring that AI agents operate within a secure and compliant framework. It includes robust access management, data encryption, and monitoring capabilities to protect sensitive information and prevent unauthorized usage. Businesses can customize AI models while maintaining strict adherence to regulatory standards, ensuring responsible AI deployment. With built-in audit trails and policy enforcement mechanisms, organizations can track AI interactions, detect anomalies, and mitigate risks effectively. This comprehensive approach makes AI adoption more scalable, accessible, and trustworthy for enterprises looking to integrate intelligent automation into their operations.", true]
    ]);
    
    // Initiatives
    const initiativesSheet = ss.getSheetByName("Initiatives");
    initiativesSheet.getRange(2, 1, 4, 6).setValues([
      [1, "Blue Lab project - Apigee Migration Accelerator", "Apigee Migration Accelerator solutions with a focus on microservices architecture.", "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQY6ifK6f2qRYyNCAPSwg9AFPN0jENtVzTRyw&s", "Project Phoenix is an ambitious initiative aimed at modernizing outdated legacy systems by leveraging GCP's robust cloud-native solutions. This project focuses on transforming monolithic architectures into highly scalable, resilient, and efficient microservices-based solutions. By utilizing Kubernetes, serverless functions, and containerized environments, organizations can achieve greater flexibility, improve operational efficiency, and reduce overall maintenance costs. Additionally, the initiative incorporates DevOps best practices, continuous integration/continuous deployment (CI/CD) pipelines, and automated monitoring to ensure seamless system performance. Through this transformation, businesses can enhance agility, accelerate innovation, and future-proof their IT infrastructure.", true],
      [2, "Blue Lab project - CICD pipeline for Apigee API Proxy deployment", "Building a unified data platform leveraging BigQuery and Looker for real-time analytics and reporting.", "https://media.licdn.com/dms/image/v2/C4E12AQGuJIc5Wsdvrg/article-cover_image-shrink_600_2000/article-cover_image-shrink_600_2000/0/1549389176142?e=2147483647&v=beta&t=L_l6_CPdqz2GSvM1S07Lys6ZTwDymiYb7hMLCxuSSnY", "The Blue Lab project - CICD pipeline for Apigee API Proxy deployment is a powerful analytics solution built on Google Cloud's BigQuery and Looker. This initiative enables enterprises to aggregate, process, and analyze vast amounts of structured and unstructured data in real-time, facilitating data-driven decision-making. By integrating AI and ML-powered analytics, the platform provides actionable insights, predictive modeling, and automated reporting capabilities. Organizations can leverage this solution to enhance customer experience, optimize business operations, and uncover hidden trends. With advanced security and governance features, businesses can maintain compliance with industry standards while maximizing the value of their data assets.", true],
      [3, "Blue Labs - The Bluepedia Page", "For updated look for project lists and teams! Reach out to your respective leads if you want to be part of any of the above.", "https://bluelabs.com/wp-content/uploads/2024/03/og-image.png", "Blue Labs - The Bluepedia Page is a cutting-edge initiative focused on enhancing cloud security using Google's industry-leading security primitives. This framework is designed to help enterprises implement end-to-end security measures, including zero-trust architectures, advanced encryption mechanisms, and threat detection systems. By incorporating IAM (Identity and Access Management), VPC Service Controls, and security policy enforcement, organizations can protect sensitive data, mitigate risks, and ensure compliance with regulatory requirements. GCP Newsletter Hosting on GCP Cloud Run also integrates AI-driven anomaly detection and automated security incident response, ensuring robust protection against evolving cyber threats.", true],
      [4, "GCP Newsletter Hosting on GCP Cloud Run", "Implementing comprehensive security measures using Google's security primitives and best practices.", "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTJkVOBDw9EN9UcJt0hCTzuM5VfRa3BePiY6g&s", "GCP Newsletter Hosting on GCP Cloud Run is a cutting-edge initiative focused on enhancing cloud security using Google's industry-leading security primitives. This framework is designed to help enterprises implement end-to-end security measures, including zero-trust architectures, advanced encryption mechanisms, and threat detection systems. By incorporating IAM (Identity and Access Management), VPC Service Controls, and security policy enforcement, organizations can protect sensitive data, mitigate risks, and ensure compliance with regulatory requirements. GCP Newsletter Hosting on GCP Cloud Run also integrates AI-driven anomaly detection and automated security incident response, ensuring robust protection against evolving cyber threats.", true]
    ]);
    
    // Security Update
    const securitySheet = ss.getSheetByName("SecurityUpdate");
    securitySheet.getRange(2, 1, 1, 6).setValues([[
      1, 
      "Critical Security Advisory", 
      "Important updates regarding Cloud Identity and Access Management. All teams must update IAM policies by March 15th to comply with new security standards.", 
      "All organizations using Cloud IAM must update their policies to include the new resource location constraints and session duration limits. These changes are required to meet updated ISO 27001 compliance standards. Failure to implement these changes by March 15th may result in service interruptions. Detailed migration guides and policy templates are available in the Security Hub.",
      "Mandatory multi-factor authentication for all admin roles\nSession duration limits of 12 hours for privileged accounts\nGeofencing requirements for sensitive data access\nEnhanced audit logging configuration requirements",
      true
    ]]);
    
    // Technical Tip
    const tipSheet = ss.getSheetByName("TechnicalTip");
    tipSheet.getRange(2, 1, 1, 3).setValues([[
      1, 
      "Optimize your Cloud Storage costs with these best practices. Implementing lifecycle policies can automatically transition objects to lower-cost storage classes or delete them when they're no longer needed.",
      true
    ]]);
    
    // Learnings
    const learningsSheet = ss.getSheetByName("Learnings");
    learningsSheet.getRange(2, 1, 5, 5).setValues([
      [1, "Enhancing Your API Lifecycle With Artificial Intelligence", "This blog will throw light on the impact of AI on the emerging trends in API Management and Governance....", "https://www.bluealtair.com/blog/enhancing-your-api-lifecycle-with-artificial-intelligence", true],
      [2, "Amplifying the Power of Google Ads with BigQuery", "Integrating Google Ads data with BigQuery, Google's managed data warehouse, enables advertisers to consolidate the data generated by Google Ads and Enterprise data sources....", "https://www.bluealtair.com/blog/amplifying-the-power-of-google-ads-with-bigquery", true],
      [3, "Redefining Connectivity with Google Cloud's Application Integration", "Google now has an answer with the release of their Application Integration iPaaS solution that addresses the future of connectivity and augments the capabilities of their Apigee API Management platform....", "https://www.bluealtair.com/blog/google-cloud-application-integration-solutions", true],
      [4, "Google Cloud Datasheet: Blue Altair", "Google Cloud services enable enterprises with all these desirable benefits in their digital excellence journey...", "https://www.bluealtair.com/hubfs/New%20Site%202023/Documents/Blue%20Altair%20-%20Google%20Cloud%20Partner%20Datasheet.pdf?hsLang=en", true],
      [5, "Apigee Datasheet: Blue Altair", "Apigee is a key driver for expediting digital transformation. Blue Altair helps clients leverage Apigee's capabilities to....", "https://www.bluealtair.com/hubfs/New%20Site%202023/Documents/Blue%20Altair%20-%20Apigee%20Partner%20Datasheet.pdf?hsLang=en", true]
    ]);
    
    // Events
    const eventsSheet = ss.getSheetByName("Events");
    eventsSheet.getRange(2, 1, 2, 5).setValues([
      [1, "Google Cloud Datasheet: Blue Altair", "Google Cloud services enable enterprises with all these desirable benefits in their digital excellence journey...", "https://www.bluealtair.com/hubfs/New%20Site%202023/Documents/Blue%20Altair%20-%20Google%20Cloud%20Partner%20Datasheet.pdf?hsLang=en", true],
      [2, "Apigee Datasheet: Blue Altair", "Apigee is a key driver for expediting digital transformation. Blue Altair helps clients leverage Apigee's capabilities to....", "https://www.bluealtair.com/hubfs/New%20Site%202023/Documents/Blue%20Altair%20-%20Apigee%20Partner%20Datasheet.pdf?hsLang=en", true]
    ]);
    
    // Certifications
    const certSheet = ss.getSheetByName("Certifications");
    certSheet.getRange(2, 1, 5, 6).setValues([
      [1, "Ayush Sagore", "Achieved Google Cloud Professional Cloud Developer certification", "2nd January 2025", "https://ca.slack-edge.com/T03K4UPRDE1-U03PG9GQ4Q2-955a3e821083-192", true],
      [2, "Ankit Agade", "Achieved Google Cloud Professional Cloud Developer certification", "3rd January 2025", "https://ca.slack-edge.com/T03K4UPRDE1-U04KS8G78M6-4758a86d2fab-512", true],
      [3, "Abdul Qudir Tinwala", "Achieved Google Cloud Professional Cloud Developer certification", "3rd January 2025", "https://ca.slack-edge.com/T03K4UPRDE1-U03MDTK8WAX-911162a3843c-192", true],
      [4, "Disha baghele", "Achieved Google Cloud Professional Cloud Developer certification", "15th February 2025", "https://ca.slack-edge.com/T03K4UPRDE1-U04JVV3PN06-30b387a17488-192", true],
      [5, "Zarqua Qazi", "Completed The \"HashiCorp Certified - Terraform-Associate", "30th March 2025", "https://ca.slack-edge.com/T03K4UPRDE1-U04MJJC13KN-8b233b596a87-192", true]
    ]);
    
    // Shoutouts
    const shoutoutSheet = ss.getSheetByName("Shoutouts");
    shoutoutSheet.getRange(2, 1, 6, 5).setValues([
      [1, "Prasad Walke", "For creating Apigee Accelerator for Apigee API Proxy and Artifact migration from One Organization (Edge/Hybrid) to another Organization.", "https://ca.slack-edge.com/T03K4UPRDE1-U03MDTMCZAP-75e827ca08b2-512", true],
      [2, "Arjun Ashok", "For helping us with initial design of the newsletter.", "https://ca.slack-edge.com/T03K4UPRDE1-U03N3N1AMPA-e20aa3911672-192", true],
      [3, "Vinay More", "for creating newsletter Framework using HTML, CSS and JS.", "https://ca.slack-edge.com/T03K4UPRDE1-U08BETX28FM-2f1b8b7271c4-512", true],
      [4, "Parthesh Patel", "for creating newsletter Framework using HTML, CSS and JS.", "https://ca.slack-edge.com/T03K4UPRDE1-U08BU14SD18-d1428cb13fdf-512", true],
      [5, "Sanket Choughule", "Creating POC for CICD pipeline building using GCP Cloud Build & Repository.", "https://ca.slack-edge.com/T03K4UPRDE1-U04PC8M8WMA-72e4be4c2ab6-512", true],
      [6, "Supriya Badgujar", "For helping BA Team members for GCP certification and notify about GCP events.", "https://ca.slack-edge.com/T03K4UPRDE1-U05JYGLEYN5-0da477cceda5-512", true]
    ]);
    
    // Next Month Focus
    const focusSheet = ss.getSheetByName("NextMonthFocus");
    focusSheet.getRange(2, 1, 2, 6).setValues([
      [1, "Cloud Native Architecture", "Deep dive into microservices and containerization strategies for modern application development.", "☁️", "cloud-icon", true],
      [2, "Security Best Practices", "Enhanced security measures and compliance updates to meet industry standards and protect sensitive data.", "🔒", "security-icon", true]
    ]);
    
    // Social Links
    const socialSheet = ss.getSheetByName("SocialLinks");
    socialSheet.getRange(2, 1, 4, 4).setValues([
      [1, "https://www.linkedin.com/company/blue-altair", "./linkedin.png", true],
      [2, "https://x.com/bluealtair1", "./twitter.png", true],
      [3, "https://www.facebook.com/bluealtair1/", "./fb.png", true],
      [4, "https://www.instagram.com/bluealtair1/#", "./ig.png", true]
    ]);
  }
  
  /**
   * Menu item to set up the spreadsheet
   */
  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Newsletter Tools')
      .addItem('Set Up Newsletter Spreadsheet', 'createNewsletterSpreadsheet')
      .addItem('Generate Newsletter JSON', 'generateNewsletterJson')
      .addToUi();
  }