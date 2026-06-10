
import React, { useEffect, useState, useCallback } from "react";
import {
  Box, Typography, Button, CircularProgress, Stack,
  Divider, LinearProgress, Paper, Collapse, FormControlLabel, Checkbox,
  Fade
} from "@mui/material";
import {
  Person as PersonIcon, Subject as SubjectIcon, CloudUpload as CloudUploadIcon,
  ErrorOutlined as ErrorIcon, Email as EmailIcon, ExpandMore as ExpandMoreIcon,
  Description as BodyIcon
} from "@mui/icons-material";

import Logo from "../../../../assets/icon-64.png";

import BookingBuilderDialog from "../dialogs/BookingBuilderDialog";
import ProfileHeader from "./ProfileHeader"; 
import parsingLoader from "../../../../assets/D_parsing-loader.webm";
import { getActiveEmailDetails } from "../../utils/officeHelper";
import { parseEmail } from "../../services/parsingService";
import { useSnackbar } from "../../context/SnackbarContext";
import { EmailDetails } from "../../types/email";
import AttachmentsList from "./AttachmentsList";
import DataRow from "./DataRow";

import ProcessingLoader from "../Loader/ProcessingLoader";
import { useNavigate } from "react-router-dom";

// Premium Glassmorphism Style
const glassStyle = {
  background: "rgba(255, 255, 255, 0.65)",
  backdropFilter: "blur(20px) saturate(180%)",
  borderRadius: "16px",
  border: "1px solid rgba(255, 255, 255, 0.5)",
  boxShadow: "0 8px 32px 0 rgba(31, 38, 135, 0.07)",
  transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
  "&:hover": {
    boxShadow: "0 12px 40px 0 rgba(31, 38, 135, 0.12)",
    transform: "translateY(-1px)"
  }
};

const Home = () => {
  const { showSnackbar } = useSnackbar();
  const [loading, setLoading] = useState(false);
  const [email, setEmail] = useState<EmailDetails | null>(null);
  const [progress, setProgress] = useState(0);


  // UI States
  const [expanded, setExpanded] = useState(false);
  const [sendAsHtml, setSendAsHtml] = useState(true);
  const [sendAsText, setSendAsText] = useState(true);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isParsing, setIsParsing] = useState(true);
  const [isProcessingEmail, setIsProcessingEmail] = useState(false);
  const [submissionError, setSubmissionError] = useState(false);
  const [loaderKey, setLoaderKey] = useState(0);

  // handle api response url 
  const [showBookingDialog, setShowBookingDialog] = useState(false);
  const [bookingUrl, setBookingUrl] = useState("");
  const navigate = useNavigate();

  const [apiErrorMessage, setApiErrorMessage] = useState<string>("");




const handleLogout = () => {
  // Add your logout logic here (e.g., clearing tokens)
  window.localStorage.removeItem("token");
  showSnackbar("Logged out successfully", "success");
  navigate("/register");
};

  // 1. REUSABLE DATA LOADER
  const loadData = useCallback(async () => {
    setIsParsing(true); // Show loader when switching emails

    // Fetch metadata (Subject, Sender, etc.)
    const data = getActiveEmailDetails();

    if (data) {
      const validIds = data.attachments.filter(a => a.isValid).map(a => a.id);
      setSelectedIds(validIds);
    }

    // Fetch body text
    if (typeof Office !== "undefined" && Office.context?.mailbox?.item) {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setEmail(data ? { ...data, bodyText: result.value } : null);
        } else {
          setEmail(data);
        }
        setIsParsing(false);
      });
    } else {
      setEmail(data);
      setTimeout(() => setIsParsing(false), 500);
    }
  }, []);


  // 2. SELECTION CHANGE EVENT HANDLER
  useEffect(() => {
    // Initial load
    loadData();

    // Check if Office is available
    if (typeof Office !== "undefined" && Office.context?.mailbox) {

      // Define the handler function
      const onItemChanged = () => {
        console.log("Email selection changed!");
        loadData();
      };

      // Add the Event Listener for ItemChanged (Required for pin-able taskpanes)
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        onItemChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to add ItemChanged handler", result.error);
          }
        }
      );

      // Cleanup: Remove handler when component unmounts
      return () => {
        Office.context.mailbox.removeHandlerAsync(
          Office.EventType.ItemChanged,
          onItemChanged
        );
      };
    }

    return undefined;
  }, [loadData]);



  const handleProcess_Email = async () => {
    setLoaderKey((prev) => prev + 1);
    setIsProcessingEmail(true);
    setSubmissionError(false);
    if (selectedIds.length === 0 && !sendAsHtml && !sendAsText) {
      showSnackbar("Please select data to process", "warning");
      setIsProcessingEmail(false);
      return;
    }

    try {
      const item = Office.context.mailbox.item;


      // 1. GET HTML BODY
      const htmlBody = await new Promise<string>((resolve) => {
        item.body.getAsync(Office.CoercionType.Html, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            let rawHtml = result.value;

            // A. Remove line breaks and tabs
            let cleanHtml = rawHtml.replace(/[\r\n\t]+/g, " ");

            // B. Ensure it has <html> and <body> tags (The API seems to prefer this)
            if (!cleanHtml.toLowerCase().includes("<html")) {
              cleanHtml = `<html><head><title>no title</title></head><body>${cleanHtml}</body></html>`;
            }

            // C. Final trim to remove leading/trailing whitespace
            resolve(cleanHtml.trim());
          } else {
            resolve("");
          }
        });
      });

      // 2. GET TEXT BODY
      const textBody = await new Promise<string>((resolve) => {
        item.body.getAsync(Office.CoercionType.Text, (result) => {
          resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
        });
      });

      // 3. MAP RECIPIENTS (Ensuring RecipientType is Integer)
      const recipients = [
        ...(Array.isArray(item.to) ? item.to : []).map((r: any) => ({
          RecipientType: 1, // To
          DisplayName: r.displayName || "",
          EmailAddress: r.emailAddress || "",
        })),
        ...(Array.isArray(item.cc) ? item.cc : []).map((r: any) => ({
          RecipientType: 2, // CC
          DisplayName: r.displayName || "",
          EmailAddress: r.emailAddress || "",
        })),
        ...(Array.isArray(item.bcc) ? item.bcc : []).map((r: any) => ({
          RecipientType: 3, // BCC
          DisplayName: r.displayName || "",
          EmailAddress: r.emailAddress || "",
        })),
      ];

      // 4. MAP ATTACHMENTS
      const selectedAttachments = email?.attachments.filter((a) => selectedIds.includes(a.id)) || [];
      const attachmentsPayload = await Promise.all(
        selectedAttachments.map(async (att, index) => {
          const base64Content = await new Promise<string>((resolve, reject) => {
            item.getAttachmentContentAsync(att.id, (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Office.js returns different structures depending on platform. 
                // Usually result.value.content is the base64 string.
                resolve(result.value.content);
              } else {
                reject(result.error);
              }
            });
          });

          return {
            AttachmentID: index,
            FileName: att.name,
            ContentType: att.contentType,
            ContentId: "",
            Content: base64Content,
          };
        })
      );

      // 5. CONSTRUCT THE emailMessage OBJECT (Mirroring Postman)
      const emailMessage = {
        Attachments: attachmentsPayload,
        Recipients: recipients,
        ReplyTo: {
          RecipientType: 1,
          DisplayName: "",
          EmailAddress: "",
        },
        Sender: {
          RecipientType: 1,
          DisplayName: email?.senderName || "",
          EmailAddress: email?.senderEmail || "",
        },
        ReceivedBy: {
          RecipientType: 1,
          DisplayName: Office.context.mailbox.userProfile.displayName || "",
          EmailAddress: Office.context.mailbox.userProfile.emailAddress || "",
        },
        Subject: email?.subject || "",
        BodyText: sendAsText ? textBody : "",
        BodyHtml: sendAsHtml ? htmlBody : "",
      };

      // 6. CALL API
      console.log("SENDING TO API:", emailMessage);
      const response = await parseEmail(emailMessage);

      if (response.isSuccess && response.data) {
        setBookingUrl(response.data);
        setShowBookingDialog(true);
        showSnackbar("Email processed successfully", "success");
      } else {
        throw new Error(response.error?.message || "Server rejected request");
      }

    } catch (err) {
      console.error("Parsing Error:", err);

      // Extract the real error message
      const msg = err?.message || "An unexpected error occurred";
      setApiErrorMessage(msg); // Store it to pass to the report page
      setSubmissionError(true);

      // Show the ACTUAL error in the snackbar
      showSnackbar(msg, "error", () => navigate("/report", { state: { error: msg } }));
    } finally {
      setIsProcessingEmail(false);
    }
  };




  if (isParsing || !email) {
    return (
      <Box
        sx={{
          height: "100vh",
          width: "100vw",
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
          background: "#ffffff",
          overflow: "hidden",
          textAlign: "center",
          position: "fixed",
          top: 0,
          left: 0,
          zIndex: 9999,
        }}
      >
        {/* Container for Video to ensure it scales */}
        <Box
          sx={{
            width: "100%",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            mb: 2
          }}
        >
         
        </Box>

        <Box sx={{ px: 4 }}>
          <Typography
            variant="h5"
            sx={{
              fontWeight: 800,
              color: "#1e293b",
              mb: 1,
              fontSize: "1.4rem" // Responsive sizing
            }}
          >
            Parsing Email
          </Typography>

          <Typography
            variant="body1"
            sx={{
              color: "#64748b",
              maxWidth: 280,
              mx: "auto",
              lineHeight: 1.5,
              fontWeight: 500,
            }}
          >
            Please wait while we extract data and attachments...
          </Typography>

          {/* Dynamic Progress - Since we don't know the exact % of the fetch, 
            we use 'indeterminate' for a smooth pulsing look */}
          <Box sx={{ width: "100%", mt: 4 }}>
            <LinearProgress
              variant="indeterminate"
              sx={{
                width: "180px",
                mx: "auto",
                height: 6,
                borderRadius: 10,
                bgcolor: "rgba(0,0,0,0.05)",
                "& .MuiLinearProgress-bar": {
                  borderRadius: 10,
                  backgroundColor: "#6aa2ba",
                },
              }}
            />
          </Box>
        </Box>
      </Box>
    );
  }
  return (
    <Box sx={{
      display: 'flex', flexDirection: 'column', height: "100vh",
      p: 2, pb: "0 !important", // Remove bottom padding to let footer sit flush
      background: "radial-gradient(circle at top left, #f8fafc, #e2e8f0)",
      boxSizing: "border-box", overflow: "hidden"
    }}>

      <ProcessingLoader
        key={loaderKey}
        open={isProcessingEmail}
      />

        <ProfileHeader
        onLogout={handleLogout}      />

        <Divider sx={{mt:1}} />


      {/* 1. SCROLLABLE CONTENT AREA */}
      <Box sx={{ flexGrow: 1, overflowY: 'auto', pr: 0.5, pb: 2, '&::-webkit-scrollbar': { width: '4px' } }}>
        <Fade in={true} timeout={800}>
          <Box>
            <Typography variant="overline" sx={{ fontWeight: 900, color: "primary.main", ml: 1, letterSpacing: 2, opacity: 0.8 }}>
              Email Summary
            </Typography>

            <Paper sx={{ ...glassStyle, p: 2, mb: 2, mt: 1 }}>
              <Stack spacing={1.8}>
                <DataRow icon={<SubjectIcon sx={{ fontSize: 18 }} />} label="Subject" value={email.subject} isTitle />
                <Divider sx={{ opacity: 0.1 }} />
                <DataRow icon={<PersonIcon sx={{ fontSize: 18 }} />} label="Origin" value={email.senderName} subValue={email.senderEmail} />
                <DataRow icon={<EmailIcon sx={{ fontSize: 18 }} />} label="Recipient" value={email.userEmail} />

                <Box sx={{ pt: 0.5 }}>
                  <Button
                    fullWidth variant="outlined" size="small"
                    onClick={() => setExpanded(!expanded)}
                    endIcon={<ExpandMoreIcon sx={{ transform: expanded ? 'rotate(180deg)' : 'rotate(0deg)', transition: '0.4s cubic-bezier(0.4, 0, 0.2, 1)' }} />}
                    sx={{ textTransform: 'none', borderRadius: '10px', fontWeight: 700, fontSize: '0.65rem', py: 0.5, borderColor: 'rgba(0,0,0,0.1)', color: 'text.secondary' }}
                  >
                    Email Body
                  </Button>
                  <Collapse in={expanded}>
                    <Paper sx={{ mt: 1.5, p: 1.5, bgcolor: 'rgba(0,0,0,0.02)', borderRadius: '12px', border: '1px dashed rgba(0,0,0,0.1)' }}>
                      <Typography variant="caption" sx={{ display: 'block', maxHeight: 90, overflowY: 'auto', fontStyle: 'italic', lineHeight: 1.6, color: 'text.secondary' }}>
                        {email.bodyText || "Analyzing payload content..."}
                      </Typography>
                      <Stack direction="row" spacing={1} sx={{ mt: 1.5 }}>
                        <FormControlLabel control={<Checkbox size="small" checked={sendAsHtml} onChange={(e) => setSendAsHtml(e.target.checked)} />} label={<Typography sx={{ fontSize: '0.65rem', fontWeight: 800 }}>HTML</Typography>} />
                        <FormControlLabel control={<Checkbox size="small" checked={sendAsText} onChange={(e) => setSendAsText(e.target.checked)} />} label={<Typography sx={{ fontSize: '0.65rem', fontWeight: 800 }}>TEXT</Typography>} />
                      </Stack>
                    </Paper>
                  </Collapse>
                </Box>
              </Stack>
            </Paper>

            <AttachmentsList
              attachments={email.attachments}
              selectedIds={selectedIds}
              onToggleOne={(id: string) => setSelectedIds(prev => prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id])}
              onSelectAll={() => {
                const validIds = email.attachments.filter(a => a.isValid).map(a => a.id);
                setSelectedIds(selectedIds.length === validIds.length ? [] : validIds);
              }}
              glassStyle={glassStyle}
            />
          </Box>
        </Fade>
      </Box>

      {/* 2. COMPACT ACTION FOOTER */}
      <Paper elevation={0} sx={{
        mx: -2, px: 2, pt: 1.5, pb: 2,
        background: "rgba(255, 255, 255, 0.85)",
        backdropFilter: "blur(25px)",
        borderTop: "1px solid rgba(0,0,0,0.06)",
        zIndex: 1000
      }}>
        {loading && (
          <Box sx={{ mb: 1.5 }}>
            <LinearProgress variant="determinate" value={progress} sx={{ height: 3, borderRadius: 10, bgcolor: 'rgba(0,0,0,0.05)' }} />
          </Box>
        )}
        <Stack spacing={1}>
          <Button
            fullWidth variant="contained" size="small" disableElevation
            onClick={handleProcess_Email} disabled={loading}
            startIcon={loading ? <CircularProgress size={14} color="inherit" /> : <CloudUploadIcon sx={{ fontSize: 16 }} />}
            sx={{
              py: 1, fontWeight: 800, textTransform: 'none', borderRadius: '10px', fontSize: '0.8rem',
              background: 'linear-gradient(45deg, #6aa2ba, #6aa2ba)',
              boxShadow: '0 2px 8px rgba(25, 118, 210, 0.2)',
              "&:hover": { background: 'linear-gradient(45deg, #6aa2ba, #6698adff)' }
            }}
          >
            {loading ? "Processing..." : `Process Email`}
          </Button>

          {submissionError && (
            <Fade in={submissionError}>
              <Button
                fullWidth variant="text" color="error" size="small"
                startIcon={<ErrorIcon sx={{ fontSize: 16 }} />}
                // PASS THE STATE HERE
                onClick={() => navigate("/report", { state: { error: apiErrorMessage } })}
                sx={{
                  textTransform: 'none', fontWeight: 700, fontSize: '0.75rem', py: 0.5,
                  bgcolor: 'rgba(211, 47, 47, 0.05)',
                  "&:hover": { bgcolor: 'rgba(211, 47, 47, 0.1)' }
                }}
              >
                Report Error
              </Button>
            </Fade>
          )}
        </Stack>
      </Paper>


      <BookingBuilderDialog
        open={showBookingDialog}
        bookingUrl={bookingUrl}
        onClose={() => setShowBookingDialog(false)}
        onOpen={() => {
          window.open(bookingUrl, "_blank");
          setShowBookingDialog(false);
        }}
      />
    </Box>
  );
};



export default Home;

