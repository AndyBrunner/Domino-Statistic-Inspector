import java.util.Calendar;
import java.util.Date;
import java.util.Hashtable;
import java.util.Properties;
import java.util.Vector;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.StringReader;
import javax.script.ScriptEngineManager;
import javax.script.ScriptEngine;
import lotus.domino.Database;
import lotus.domino.View;
import lotus.domino.Document;
import lotus.domino.DocumentCollection;
import lotus.domino.RichTextItem;
import lotus.domino.RichTextStyle;
import lotus.domino.DateTime;
import lotus.domino.NotesException;

public class StatInspector extends JAddinThread
{
	// Instance variables
	String							xVersionNumber				= "0.8.0";			//TODO: Set version number, also in StatInspector.ntf
	boolean							xCleanupDone				= false;
	String							xDatabaseName				= "StatInspector.nsf";
	String							xCurrentServerCanonical		= null;
	String							xCurrentServerAbbreviated	= null;
	Database						xNotesDatabase				= null;
	boolean							xDeleteAlertDocuments		= true;
	
	// Main entry point
	public void addinStart()
	{
		// Welcome message
		logMessage("Domino Server Statistic Inspector Version " + xVersionNumber + " (Freeware)");
		
		// Set addin state
		setAddinState("Initialization in progress");
		
		// Save passed command line parameter (database name)
		if (getAddinParameters() != null)
			xDatabaseName = getAddinParameters();
		
		// Get current server name
		try {
			xCurrentServerCanonical		= getDominoSession().getServerName();
			xCurrentServerAbbreviated	= getDominoSession().createName(xCurrentServerCanonical).getAbbreviated();
		} catch (NotesException e) {
			logMessage("Error: Unable to get current Domino server name: " + e.text);
			cleanup();
			return;
		}
		
		// Open StatInspector database
		try {
			logDebug("Open database " + xDatabaseName);
			xNotesDatabase = getDominoSession().getDatabase("", xDatabaseName);
			
			if (!xNotesDatabase.isOpen()) {
				logMessage("Error: Unable to open database " + xDatabaseName);
				cleanup();
				return;
			}
			
			// Get full path and filename (with extension)
			xDatabaseName = xNotesDatabase.getFilePath();
			
		} catch (NotesException e) {
			logMessage("Error: Unable to open database " + xDatabaseName + ": " + e.text);
			cleanup();
			return;
		}

		// Loop for each time interval
		View							notesConfigurationView		= null;
		View							notesProbeView				= null;
		DocumentCollection				notesDocumentCollection		= null;
		Document						notesProbeDocument			= null;
		Document						notesConfigurationDocument	= null;
		Document						notesAlertDocument			= null;
		Document						notesMemoDocument			= null;
		RichTextItem					notesRichTextItem			= null;
		Vector<?>						statisticNames				= null;
		Vector<?>						collectFromServers			= null;
		Vector<?>						alertAddresses				= null;	
		Hashtable<String, Properties>	lastRunStatistics			= new Hashtable<String, Properties>();
		long							delayTimeMinutes			= 0L;
		long							alertExpirationDays			= 0L;
		int								alertDocumentsCreated		= 0;
		
		while (true) {

			// Pseudo loop
			while (true) {

				delayTimeMinutes		= 30L;
				alertDocumentsCreated	= 0;

				// Get the configuration document for this collecting server
				try {
					logDebug("Get configuration document");

					// Set addin state
					setAddinState("Reading configuration document");

					notesConfigurationView	= xNotesDatabase.getView("Configurations");
					notesDocumentCollection	= notesConfigurationView.getAllDocumentsByKey(xCurrentServerAbbreviated);

					if (notesDocumentCollection.getCount() == 0) {
						logMessage("Error: No configuration document found for current server in " + xDatabaseName);
						break;
					}

					if (notesDocumentCollection.getCount() > 1) {
						logMessage("Error: More than one configuration document found for current server in " + xDatabaseName);
						break;
					}

					notesConfigurationDocument = notesDocumentCollection.getFirstDocument();

					collectFromServers	= notesConfigurationDocument.getItemValue("CollectFromServers");
					statisticNames		= notesConfigurationDocument.getItemValue("StatisticNames");
					delayTimeMinutes	= notesConfigurationDocument.getItemValueInteger("IntervalMinutes");
					alertExpirationDays	= notesConfigurationDocument.getItemValueInteger("AlertExpirationDays");
					alertAddresses		= notesConfigurationDocument.getItemValue("AlertAddresses");

					if (collectFromServers != null) {
						if (collectFromServers.size() == 0)
							collectFromServers = null;
					}

					if (statisticNames != null) {
						if (statisticNames.size() == 0)
							statisticNames = null;
					}

					if (alertAddresses != null) {
						if (alertAddresses.size() == 0)
							alertAddresses = null;
					}

					if (delayTimeMinutes == 0L)
						delayTimeMinutes = 30L;

				} catch (NotesException e) {
					logMessage("Error: Unable to read configuration document in " + xDatabaseName + ": " + e.text);
					break;
				}

				// Delete expired alert documents (either at startup or after midnight)
				if (xDeleteAlertDocuments) {

					if (alertExpirationDays != 0) {

						Calendar expirationDate = Calendar.getInstance();
						expirationDate.add(Calendar.DAY_OF_MONTH, (int) -alertExpirationDays);

						logDebug("Deleting all alert documents before " + expirationDate.getTime());

						deleteExpiredAlertDocuments(expirationDate.getTime());
					}

					xDeleteAlertDocuments = false;
				}

				// Open probe view
				try {
					notesProbeView = xNotesDatabase.getView("(Probes enabled)");
				} catch (NotesException e) {
					logMessage("Error: Unable to open view '(Probes enabled)' in " + xDatabaseName + ": " + e.text);
					break;
				}

				// Load JavaScript engine
				ScriptEngine javaScriptEngine = new ScriptEngineManager().getEngineByName("JavaScript");

				// Loop thru all defined servers
				String		serverName				= null;
				String		serverNameAbbreviated	= null;
				Object		scriptResult			= null;
				Vector<?>	probeServers			= null;
				String		probeScript				= null;
				String		probeMessage			= null;
				String		probeID					= null;

				for (int sIndex = 0; sIndex < collectFromServers.size(); sIndex++) {

					serverName = (String) collectFromServers.get(sIndex);
					try {
						serverNameAbbreviated = getDominoSession().createName(serverName).getAbbreviated();
					} catch (NotesException e) {
						logMessage("Error: Unable to get abbreviated server name: " + e.text);
						serverNameAbbreviated = serverName;
					}

					logDebug("Processing server " + serverName);

					// Set addin state
					setAddinState("Analyzing statistics from server " + serverNameAbbreviated);

					// Read thru all probe documents
					logDebug("Get probe documents");

					// Get server statistics
					Properties serverStatistics = getStatistics(serverName, statisticNames);
					if (serverStatistics == null)
						continue;

					try {
						notesProbeDocument = notesProbeView.getFirstDocument();

						// Loop thru all probe documents
						int				alertCount		= 0;
						Vector<String>	messagesText	= new Vector<String>(0,1);

						while (notesProbeDocument != null) {

							// Get all fields
							probeServers	= notesProbeDocument.getItemValue("Servers");
							probeScript		= notesProbeDocument.getItemValueString("JavaScript");
							probeMessage	= notesProbeDocument.getItemValueString("Message");
							probeID			= notesProbeDocument.getItemValueString("ProbeID");

							if (probeServers != null) {
								if (probeServers.size() == 0)
									probeServers = null;
							}

							logDebug("Processing probe document " + probeID);

							while (true) {

								String	probeComposedMessage = null;

								// Check if server list defined and server name included
								if (probeServers != null) {

									boolean serverFound = false;

									for (int index = 0; index < probeServers.size(); index++) {

										if (((String) probeServers.get(index)).equalsIgnoreCase(xCurrentServerCanonical)) {
											serverFound = true;
											break;
										}
									}

									if (!serverFound)
										break;
								}

								// Replace all statistic names with the statistic value and invoke the JavaScript engine
								try {

									probeScript = replaceStatisticValues(probeScript, serverStatistics, statisticNames, serverNameAbbreviated, lastRunStatistics);

									// Ignore probe if statistic not found
									if (probeScript == null)
										break;

									scriptResult	= javaScriptEngine.eval("function execute() { " + probeScript + " } ; execute()");
									logDebug("JavaScript engine returned <" + scriptResult + '>');

									if (!(scriptResult instanceof Boolean))
										break;

									if (!((Boolean) scriptResult))
										break;	

								} catch (Exception e) {
									probeComposedMessage = "Error: Syntax error in JavaScript code: " + e.getMessage();
								}

								// Increase alert counter
								alertCount++;

								// Create new alert document if necessary
								if (notesAlertDocument == null) {

									logDebug("Create new alert document");

									notesAlertDocument = xNotesDatabase.createDocument();
									notesAlertDocument.replaceItemValue("Form", "Alert");
									notesAlertDocument.replaceItemValue("TimeStamp", getDominoSession().createDateTime(new Date()));
									notesAlertDocument.replaceItemValue("ServerName", xCurrentServerCanonical);
									notesRichTextItem = notesAlertDocument.createRichTextItem("Messages");				
								}

								// Replace statistic name with statistic value in message
								if (probeComposedMessage == null) {

									try {
										probeComposedMessage = replaceStatisticValues(probeMessage, serverStatistics, statisticNames, serverNameAbbreviated, lastRunStatistics);

										// Set original alert message if statistic not found
										if (probeComposedMessage == null)
											probeComposedMessage = probeMessage;

									} catch (Exception e) {
										probeComposedMessage = "Error: Syntax error in alert message: " + e.getMessage();
									}
								}

								// Append document link and message to alert document
								if (alertCount > 1)
									notesRichTextItem.addNewLine();

								notesRichTextItem.appendDocLink(notesProbeDocument);
								notesRichTextItem.appendText(' ' + probeComposedMessage);

								messagesText.add(probeComposedMessage); 

								break;
							}

							// Get next document in view
							Document notesDocumentTemp = notesProbeDocument;
							notesProbeDocument = notesProbeView.getNextDocument(notesProbeDocument);
							notesDocumentTemp.recycle();
						}

						// Save alert document if created
						if (notesAlertDocument != null) {

							logDebug("Saving alert document");

							notesAlertDocument.replaceItemValue("AlertCount", new Integer(alertCount));
							notesAlertDocument.replaceItemValue("MessagesText", messagesText);

							notesAlertDocument.computeWithForm(false, false);
							notesAlertDocument.save(true);

							// Mail document if recipient defined
							if (alertAddresses != null)
								sendAlertMessage(serverName, alertAddresses, notesRichTextItem);

							notesRichTextItem.recycle();
							notesRichTextItem = null;

							notesAlertDocument.recycle();
							notesAlertDocument = null;

							alertDocumentsCreated++;
						}

					} catch (NotesException e) {
						logMessage("Error: Unable to process probe documents in " + xDatabaseName + ": " + e.text);
						return;
					}

					// Save current server statistics for next time interval
					lastRunStatistics.put(serverNameAbbreviated, serverStatistics);
				}

				if (alertDocumentsCreated > 0)
					logMessage("Created " + alertDocumentsCreated + " alert document" + ((alertDocumentsCreated > 1) ? 's' : "") + " in " + xDatabaseName);

				break;
			}

			// Recycle all Domino objects used during this interval
			try {
				logDebug("Recycle all Domino objects");

				if (notesRichTextItem != null) {
					notesRichTextItem.recycle();
					notesRichTextItem = null;
				}

				if (notesConfigurationDocument != null) {
					notesConfigurationDocument.recycle();
					notesConfigurationDocument = null;
				}

				if (notesProbeDocument != null) {
					notesProbeDocument.recycle();
					notesProbeDocument = null;
				}

				if (notesAlertDocument != null) {
					notesAlertDocument.recycle();
					notesAlertDocument = null;
				}

				if (notesMemoDocument != null) {
					notesMemoDocument.recycle();
					notesMemoDocument = null;
				}

				if (notesDocumentCollection != null) {
					notesDocumentCollection.recycle();
					notesDocumentCollection = null;
				}

				if (notesConfigurationView != null) {
					notesConfigurationView.recycle();
					notesConfigurationView = null;
				}

				if (notesProbeView != null) {
					notesProbeView.recycle();
					notesProbeView = null;
				}

			} catch (Exception e) {
				logMessage("Error: Unable to recycle Domino objects: " + e.getMessage());
			}

			// Set addin state
			setAddinState(null);

			// Delay the specified number of minutes
			waitMilliSeconds(delayTimeMinutes * 60000L);
		}
	}

	// Add-In received "Quit" or "Exit" command
	public void addinStop()
	{
		logMessage("Termination in progress");
		cleanup();
		addinTerminate();
	}
	
	// Addin received notification of midnight
	public void addinNextDay() {
		// Set flag to delete expired alert documents
		xDeleteAlertDocuments = true;
	}
	
	// Addin received command thru "Tell" command.
	public void addinCommand(String pCommand)
	{
		logDebug("User command <" + pCommand + "> received");
		logMessage("Error: This addin does not support any command except 'Quit'");
		logMessage("Use 'Tell StatInspector Help!' for help on the special commands of the JAddin framework");
	}
	
	/** 
	 * Get all Domino server statistics and create property object
	 * 
	 * @param  serverName Domino server to be collected
	 * @return Property All Domino server statistics
	 */
	Properties getStatistics(String pServerName, Vector<?> pStatistics)
	{
		Properties		statValues		= new Properties();
		BufferedReader	reader			= null;
		String			statisticLine	= null;
		String			statisticValues = null;

		logDebug("Method getStatistics() called");

		// Read all defined statistic values
		for (int sIndex = 0; sIndex < pStatistics.size(); sIndex++) {

			logDebug("Reading Domino server statistics <" + pStatistics.get(sIndex) + "> from server " + pServerName);
			
			// Get statistic values
			try {
				// Undocumented: If the server command starts with "!", the response do not show up in log.nsf
				statisticValues = getDominoSession().sendConsoleCommand(pServerName, "!Show Statistic " + pStatistics.get(sIndex));
			} catch (NotesException e) {
				logMessage("Error: Unable to retrieve statistics from Domino server " + pServerName + ": " + e.text);
				return null;	
			}
			
			// Read the string with statistics per line and add it to the property 
			try {
				
				reader = new BufferedReader(new StringReader(statisticValues));
				
				while ((statisticLine = reader.readLine()) != null) {
					
					statisticLine = statisticLine.trim();
					
					if (statisticLine.length() == 0)
						continue;
					
					int index = statisticLine.indexOf('=');
					
					// Skip invalid entry
					if (index == -1)
						continue;	

					// Add the statistic name and value to the property
					statValues.setProperty(statisticLine.substring(0, index).trim().toUpperCase(), statisticLine.substring(index + 1).trim());
				}
				
				reader.close();
				
			} catch (IOException e) {
				logMessage("Error: Unable to interpret Domino server statistics <" + pStatistics.get(sIndex) + ">: " + e.getMessage());
				return null;
			}
		}
		
		// Add timestamp (number of seconds) to the properties
		statValues.setProperty("$STATINSPECTOR.TIMESTAMP", (String) Long.toString(System.currentTimeMillis() / 1000L));
		
		logDebug("Number of Domino server statistics read: " + statValues.size());
		
		return statValues;
	}

	/** 
	 * Replace all statistic names in brackets with the statistic value.
	 * 
	 * @param	psourceString	String to be changed
	 * @param	pstatistics	Properties with the statistics
	 * @return	Changed string or null if errors
	 * @throws	Exception If syntax errors
	 */
	String replaceStatisticValues(String pSourceString, Properties pStatistics, Vector<?> pConfiguredStatistics, String pServerNameAbbreviated, Hashtable<String, Properties> pSavedStatistics) throws Exception {
		
		String	statisticVariable		= null;
		String	statisticName			= null;
		String	statisticValue			= null;
		int		startPosition			= 0;
		int		endPosition				= 0;
		
		logDebug("Method replaceStatisticValues() called");
		logDebug("Input string: <" + pSourceString + '>');
		
		// Loop thru the source string
		while ((startPosition = pSourceString.indexOf('[')) != -1) {
			
			endPosition = pSourceString.indexOf(']');
			
			if (endPosition == -1)
				throw new Exception("Missing close bracket character");
			
			if (endPosition < startPosition)
				throw new Exception("Missing open bracket character");
		
			statisticVariable	= pSourceString.substring(startPosition, endPosition);
			statisticName		= statisticVariable.substring(1, statisticVariable.length()).trim().toUpperCase();
			statisticValue		= "";
			
			// Replace special variable names
			if (statisticName.equals("$INTERVAL")) {
				
				Properties previousStatistics = pSavedStatistics.get(pServerNameAbbreviated);
				
				// Exit if no previous statistics
				if (previousStatistics == null) {
					pSourceString = null;
					break;						
				}
								
				long currentValue	= Long.parseLong((String) pStatistics.getProperty("$STATINSPECTOR.TIMESTAMP", "0"));
				long previousValue	= Long.parseLong((String) previousStatistics.getProperty("$STATINSPECTOR.TIMESTAMP", "0"));
				
				// Test if server statistics were reset with the console command "Set Statistics xxx"
				if (currentValue < previousValue) {
					pSourceString = null;
					break;
				}
				
				statisticValue	= Long.toString(currentValue - previousValue).trim();
			}
						
			if (statisticName.equals("$SERVERNAME"))
				statisticValue = pServerNameAbbreviated;
			
			// Replace the statistic name with the statistic value or with the difference between current and last statistic value
			if (statisticValue.equals("")) {
				
				if (statisticName.startsWith("+")) {
					
					Properties previousStatistics = pSavedStatistics.get(pServerNameAbbreviated);
					
					// Exit if no previous statistics
					if (previousStatistics == null) {
						pSourceString = null;
						break;						
					}
															
					statisticName		 	= statisticName.substring(1).trim();
					statisticValue			= pStatistics.getProperty(statisticName, "").trim();
					String lastValue		= previousStatistics.getProperty(statisticName, "");
					
					// Replace value with difference between last and current value
					if (!lastValue.equals("")) {
						
						try {
							// Subtract previous value from current value and clear fraction part (e.g. 1.0 > 1)
							double doubleValue = Double.parseDouble(statisticValue) - Double.parseDouble(lastValue);
							
							if ((doubleValue % 1) > 0)
								statisticValue = Double.toString(doubleValue);
							else
								statisticValue = Integer.toString((int) doubleValue);
														
						} catch (Exception e) {
							throw new Exception("Referenced non-numeric incremental statistic name");
						}
					}
				} else
					statisticValue = pStatistics.getProperty(statisticName, "").trim();
			}
			
			logDebug("Variable name <" + statisticName + "> replaced with <" + statisticValue + '>');
			
			// Statistic name not found - check if it was selected in configuration document
			if (statisticValue.equals("")) {
				
				String	statName	= null;
				boolean statFound	= false;
				
				for (int index = 0; index < pConfiguredStatistics.size(); index++) {
					
					statName = ((String) pConfiguredStatistics.get(index)).toUpperCase() + '.';
					
					if (statisticName.toUpperCase().startsWith(statName)) {
						statFound = true;
						break;
					}
				}
				
				if (!statFound)
					throw new Exception("Referenced statistic <" + statisticName + "> not selected in configuration document");
				
				// Statistic name configured but not found
				pSourceString = null;
				break;
			}
			
			// Add string delimiters
			if (!statisticValue.equals("")) {
						
				// Check if numeric
				boolean stringIsNumeric = false;
				
				try {
					Double.parseDouble(statisticValue);
					stringIsNumeric = true;
				} catch (Exception e) {
					stringIsNumeric = false;
				}
				
				if (!stringIsNumeric)
					statisticValue = "'" + statisticValue + "'";
			}
						
			// Replace the variable
			pSourceString = pSourceString.substring(0, startPosition) + statisticValue + pSourceString.substring(endPosition + 1);
		}

		logDebug("Output string: <" + pSourceString + '>');
		return pSourceString;		
	}
	
	/**
	 * Delete expired alert documents
	 * 
	 * @param pExpireDate	Remove alert documents oder than this date.
	 * 			 
	 */
	public void deleteExpiredAlertDocuments(Date pExpireDate) {
		
		View				alertView			= null;
		DocumentCollection	alertCollection		= null;
		Document			alertDocument		= null;
		Document			alertDocumentTemp	= null;
		Date				alertDateTime		= null;
		Vector<?>			documentTimestamp	= new Vector<DateTime>(0,1);
		int					alertDeletions		= 0;
		
		logDebug("Method deleteExpiredAlertDocuments() called");
		
		try {
			alertView		= xNotesDatabase.getView("(Alerts)");
			alertCollection = alertView.getAllDocumentsByKey(xCurrentServerAbbreviated);
			
			logDebug("Number of alert documents for this server: " + alertCollection.getCount());
			
			alertDocument = alertCollection.getFirstDocument();
			
			while (alertDocument != null) {
				
				documentTimestamp	= alertDocument.getItemValue("Timestamp");
				alertDateTime		= ((DateTime) documentTimestamp.get(0)).toJavaDate();
			
				// Get next document in view and delete current document if expired
				alertDocumentTemp	= alertCollection.getNextDocument();
				
				if (alertDateTime.before(pExpireDate)) {
					logDebug("Delete alert document");
					alertDocument.remove(true);
					alertDeletions++;
				}
				
				alertDocument.recycle();
				alertDocument = alertDocumentTemp;
			}
			
			if (alertDeletions > 0)
				logMessage("Deleted " + alertDeletions + " expired alert document" + ((alertDeletions > 1) ? 's' : ""));
						
		} catch (NotesException e) {
			logMessage("Error: Unable to delete expired alert documents: " + e.text);			
		} finally {
			
			try {
				
				if (alertDocumentTemp != null) {
					alertDocumentTemp.recycle();
					alertDocumentTemp = null;
				}
				
				if (alertDocument != null) {
					alertDocument.recycle();
					alertDocument = null;
				}
				
				if (alertCollection != null) {
					alertCollection.recycle();
					alertCollection = null;
				}
					
				if (alertView != null) {
					alertView.recycle();
					alertView = null;
				}
				
			} catch (NotesException e) {
				logMessage("Error: Unable to free Domino resources: " + e.text);
			}
		}
	}
	
	/**
	 * Create and send a message with the alert information.
	 * 
	 * @param	pMessageServer The Domino server for which the alert is created
	 * @param	pMessageTo Recipient name
	 * @param	pMessageBody Body of message
	 * 
	 * @return	Success or failure indicator 			 
	 */
	public boolean sendAlertMessage(String pMessageServer, Vector<?> pMessageTo, RichTextItem pMessageBody) {
		
		// Variables
		Database		mailDatabase		= null;
		Document		mailDocument		= null;
		RichTextItem	mailRTItem			= null;
		DateTime		mailDateTime		= null;
		RichTextStyle	mailRTStyle			= null; 

		logDebug("Method sendAlertMessage() called");
				
		try {
			// Open the routers mailbox
			mailDatabase = getDominoSession().getDatabase(null, "mail.box");
			
			if (!mailDatabase.isOpen()) {
				logMessage("Error: Unable to open the Domino database mail.box");
				
				if (mailDatabase != null)
					mailDatabase.recycle();
				
				return false;
			}
				
			// Create mail message
			mailDocument = mailDatabase.createDocument();

			// Set header fields
			String senderAddress = "StatInspector@" + getDominoSession().getEnvironmentString("Domain", true);
			
			mailDocument.replaceItemValue("Form", "Memo");
			mailDocument.replaceItemValue("Principal", senderAddress);
			mailDocument.replaceItemValue("SendTo", pMessageTo);
			mailDocument.replaceItemValue("Subject", "Domino Server Statistic Inspector Alert");
			
			// Set current date and time in field <PostedDate>
			mailDateTime = getDominoSession().createDateTime("Today");
			mailDateTime.setNow();
			mailDocument.replaceItemValue("PostedDate", mailDateTime);

			mailRTItem = mailDocument.createRichTextItem("Body");
			mailRTItem.appendText("Dear Domino Administrator:");
			mailRTItem.addNewLine();
			mailRTItem.addNewLine();
			mailRTItem.appendText("The Domino Server Statistic Inspector created the following alert document for the server " + getDominoSession().createName(pMessageServer).getAbbreviated() + ':');
			mailRTItem.addNewLine();
			mailRTItem.addNewLine();
			mailRTItem.appendRTItem(pMessageBody);
			mailRTItem.addNewLine();
			mailRTItem.addNewLine();
			mailRTItem.appendText("You may open the ");
			mailRTItem.appendDocLink(xNotesDatabase);
			mailRTItem.appendText(" Statistic Inspector database on the server to view all alerts or you can use the document links above to open the probe document which triggered the alert.");
			mailRTItem.addNewLine();
			mailRTItem.addNewLine();
			mailRTItem.appendText("Regards,");
			mailRTItem.addNewLine();
			mailRTItem.appendText("Domino Server Statistics Inspector (Freeware)");
			mailRTItem.addNewLine();
			mailRTItem.addNewLine();
			mailRTItem.appendText("http://abdata.ch/StatInspector");
			mailRTItem.addNewLine();
						
			mailDocument.send();
			
		} catch (NotesException e) {
			logMessage("Error: Unable to send alert message: " + e.text);
			return false;
		} finally {
			// Free the objects
			try {
				if (mailDateTime != null)
					mailDateTime.recycle();
				
				if (mailRTStyle != null)
					mailRTStyle.recycle();
				
				if (mailRTItem != null)
					mailRTItem.recycle();
			
				if (mailDocument != null)
					mailDocument.recycle();
			
				if (mailDatabase != null)
					mailDatabase.recycle();
				
			} catch (NotesException e) {
				logMessage("Error: Unable to free the Domino resources: " + e.text);
				return false;
			}
		}
		
		logDebug("Message successfully created in <mail.box>");
		
		// Return success to the caller
		return true;
	}
	
	/**
	 * Performs all necessary cleanup tasks. 
	 */
	void cleanup() {
		
		logDebug("Method cleanup() called");
		
		// Check if cleanup already done
		if (xCleanupDone)
			return;

		try {
			logDebug("Freeing the Domino resources");
									
			if (xNotesDatabase != null) {
				xNotesDatabase.recycle();
				xNotesDatabase = null;
			}
						
		} catch (Exception e) {
			logMessage("Error: Cleanup processing failed: " + e.getMessage());
		}
		
		xCleanupDone = true;
	}
}