import * as React from "react";
import { useState } from "react";
import { Button, Field, tokens, makeStyles } from "@fluentui/react-components";

/* global HTMLTextAreaElement */

interface TextInsertionProps {
  insertText: (text: string) => void;
}

interface Attachment {
  id: string;
  name: string;
  // content: string;
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
 * TextInsertion component is a React functional component that handles the fetching and downloading of email attachments.
 * 
 * @component
 * @param {TextInsertionProps} props - The props for the TextInsertion component.
 * 
 * @returns {JSX.Element} The rendered component.
 * 
 * @example
 * <TextInsertion />
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
const TextInsertion: React.FC<TextInsertionProps> = () => {
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [isAttachmentFetched, setIsAttachmentFetched] = useState<boolean | null>(null);

  const [isSubmissionInitiated, setIsSubmissionInitiated] = useState<boolean>(false);

  const fetchAttachments = async () => {

    try {
      const names: Attachment[] = [];
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
    } catch (error) {
      console.error("Error fetching attachments", error);
    }
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      {attachments.length === 0 && (
        <>
          <Field className={styles.instructions}>
            Please provide your consent by clicking below button to extract MetaData from respective email.
          </Field>
          <Button appearance="primary" disabled={attachments.length > 0} size="medium" onClick={fetchAttachments}>
            Fetch MetaData
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
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>Sender:</strong> {Office.context.mailbox.item.sender.emailAddress}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>To:</strong> {Office.context.mailbox.item.to[0].emailAddress}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>Subject:</strong> {Office.context.mailbox.item.subject}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>ConversationId:</strong> {Office.context.mailbox.item.conversationId}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>ItemId:</strong> {Office.context.mailbox.item.itemId}</span>
          <span style={{ marginLeft: '40px', width: '100%', display: 'block' }}> <strong>Received TimeStamp:</strong> {Office.context.mailbox.item.dateTimeCreated.getTime().toString()}</span>

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

export default TextInsertion;
