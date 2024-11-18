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

const TextInsertion: React.FC<TextInsertionProps> = () => {
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [isAttachmentFetched, setIsAttachmentFetched] = useState<boolean | null>(null);

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
          } else {
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

      <Field className={styles.instructions}>
        Please click below "fetch Attachment" button to extract attachments from Email
      </Field>
      <Button appearance="primary" disabled={attachments.length > 0} size="medium" onClick={fetchAttachments}>
        Fetch Attachments
      </Button>

      {isAttachmentFetched != null && attachments.length === 0 && isAttachmentFetched == false && (
        <Field className={styles.instructions} style={{ color: 'red' }}>
          No attachments found in this email
        </Field>
      )}

      {attachments.length > 0 && (
        <>
          <h4 style={{ margin: 0, marginTop: '10px', marginLeft: '40px', width: '100%', display: 'block' }}>List of Attachments:</h4>
          <ul style={{ marginTop: 0 }}>
            {attachments.map((attachment) => (
              <li key={attachment.id}>{attachment.name}</li>
            ))}
          </ul>

          <Field className={styles.instructions}>
            Submit above attachments to SmartDoc Assistant for Automated Email Scanning and Processing
          </Field>

          <Button appearance="primary" size="large" onClick={downloadAttachments}>
            Submit
          </Button>
        </>

      )}

    </div>
  );
};

export default TextInsertion;
