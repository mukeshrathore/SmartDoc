import * as React from "react";
import { useState } from "react";
import { Button, Field, tokens, makeStyles } from "@fluentui/react-components";

/* global HTMLTextAreaElement */

interface processEmailDataProps {
  insertText: (text: string) => void;
}

interface Attachment {
  id: string;
  name: string;
  // content: string;
}
interface metadata {
  sender: string;
  to: string;
  subject: string;
  conversationId: string;
  itemId: string;
  timeStamp: string;
}

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
    marginLeft: "20px",
    marginRight: "20px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
  attachmentList: {
    marginTop: "20px",
    listStyleType: "none",
    padding: 0,
  },
  attachmentItem: {
    marginBottom: "10px",
  },
});

/**
 * processEmailData component is a React functional component that handles the fetching and downloading of email attachments.
 * 
 * @component
 * @param {processEmailDataProps} props - The props for the processEmailData component.
 * 
 * @returns {JSX.Element} The rendered component.
 * 
 * @example
 * <processEmailData />
 * 
 * @remarks
 * This component uses the Office JavaScript API to interact with email attachments in Outlook.
 * It provides functionality to fetch attachments from an email and download them.
 * 
 * @state {Attachment[]} attachments - The list of attachments fetched from the email.
 * @state {boolean | null} isAttachmentFetched - A flag indicating whether the attachments have been fetched.
 * @state {boolean} isSubmissionInitiated - A flag indicating whether the submission of attachments has been initiated.
 * 
 * @function fetchAttachments - Fetches the attachments from the email and updates the state.
 * @function downloadAttachments - Downloads the fetched attachments and updates the state.
 * 
 * @requires Office JavaScript API
 * 
 * @styles
 * - textPromptAndInsertion: The main container style.
 * - instructions: The style for the instruction text.
 * 
 * @dependencies
 * - React
 * - Office JavaScript API
 */
const processEmailData: React.FC<processEmailDataProps> = () => {
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [isAttachmentFetched, setIsAttachmentFetched] = useState<boolean | null>(null);

  const [isSubmissionInitiated, setIsSubmissionInitiated] = useState<boolean>(false);
  const [isConsent, setIsConsent] = useState<boolean>(false);
  const [metaData, setMetaData] = useState(null);

  // call a function to get metadata
  const getMetadata = async () => {
    // Simulate a call to get metadata
    setMetaData({
      sender: Office.context.mailbox.item.sender.emailAddress.toString(),
      to: Office.context.mailbox.item.to[0].emailAddress.toString(),
      subject: Office.context.mailbox.item.subject.toString(),
      conversationId: Office.context.mailbox.item.conversationId.toString(),
      itemId: Office.context.mailbox.item.itemId.toString(),
      timeStamp: Office.context.mailbox.item.dateTimeCreated.getTime().toString()
    });
    fetchAttachmentMetadata();
  };

  const fetchAttachmentMetadata = async () => {
    try {
      const names: Attachment[] = [];
      // fetching attachments names from email
      Office.context.mailbox.item.attachments.forEach(async (attachment) => {
        names.push({
          id: attachment.id,
          name: attachment.name,
        });
        setAttachments(names);
      });
      setIsAttachmentFetched(attachments.length > 0);

    } catch (error) {
      setIsAttachmentFetched(false);
      console.error('Error fetching attachments:', error);
    }
  };

  const downloadAttachments = async () => {
    try {
      Office.context.mailbox.item.attachments.forEach(async (attachment) => {
        const attachmentId = attachment.id;
        Office?.context?.mailbox?.item?.getAttachmentContentAsync(attachmentId, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const content = result.value.content;
            const contentType = result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64 ? 'application/octet-stream' : attachment.contentType;

            let blob;
            if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
              const byteCharacters = atob(content);
              const byteNumbers = new Array(byteCharacters.length);
              for (let i = 0; i < byteCharacters.length; i++) {
                byteNumbers[i] = byteCharacters.charCodeAt(i);
              }
              const byteArray = new Uint8Array(byteNumbers);
              blob = new Blob([byteArray], { type: contentType });
            } else {
              blob = new Blob([content], { type: contentType });
            }

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = attachment.name;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setIsSubmissionInitiated(true);
          } else {
            setIsSubmissionInitiated(false);
            console.error("Error fetching attachment content", result.error);
          }
        });
      });

      // download metadata
      downloadMetaData(metaData);

    } catch (error) {
      console.error("Error fetching attachments", error);
    }
  };

  const downloadMetaData = (metaData: any) => {
    const jsonString = JSON.stringify(metaData); // Formatted JSON
    const blob = new Blob([jsonString], { type: 'text/plain' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'metadata.txt';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      {attachments.length === 0 && (
        <>
          <label>
            <input
              type="checkbox"
              onChange={(e) => setIsConsent(e.target.checked)}
            />
            I consent to extract MetaData from this email.
          </label>
          {/* <Field className={styles.instructions}>
            Please provide your consent by clicking below button to extract MetaData from respective email.
          </Field> */}
          isConsent:{isConsent}
          <Button appearance="primary" disabled={!isConsent && attachments.length > 0} size="medium" onClick={getMetadata}>
            Proceed
          </Button>
        </>
      )}

      {isAttachmentFetched != null && attachments.length === 0 && isAttachmentFetched == false && (
        <Field className={styles.instructions} style={{ color: 'red' }}>
          No attachments found in this email
        </Field>
      )}

      {attachments.length > 0 && !isSubmissionInitiated && (
        <>
          <strong style={{ marginLeft: '40px', width: '100%', display: 'block' }}>MetaData:</strong>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>Sender:</strong> {metaData?.sender}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>To:</strong> {metaData?.to}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>Subject:</strong> {metaData?.subject}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>ConversationId:</strong> {metaData?.conversationId}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>ItemId:</strong> {metaData?.itemId}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>Received TimeStamp:</strong> {metaData?.timeStamp}</span>

          <h4 style={{ margin: 0, marginTop: '10px', marginLeft: '40px', width: '100%', display: 'block' }}>Attachments:</h4>
          <ul style={{ marginTop: 0 }}>
            {attachments.map((attachment) => (
              <li key={attachment.id}>{attachment.name}</li>
            ))}
          </ul>

          <Field className={styles.instructions}>
            Submit above Details to SmartDoc Assist for Automated Email Scanning and Processing
          </Field>

          <Button appearance="primary" size="large" onClick={downloadAttachments}>
            Submit
          </Button>
        </>
      )}

      {isSubmissionInitiated && (
        <Field className={styles.instructions} style={{ color: 'green' }}>
          Successfully forwarded MetaData along with attachments to SmartDoc Assist for further processing.
        </Field>
      )}

    </div>
  );
};

export default processEmailData;


