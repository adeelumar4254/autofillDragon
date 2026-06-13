
import React, { useState } from "react";
import {
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Typography,
  IconButton,
  Box,
  Tooltip,
} from "@mui/material";

interface BookingBuilderDialogProps {
  open: boolean;
  bookingUrl: string;
  onClose: () => void;
  onOpen: () => void;
}

const BookingBuilderDialog: React.FC<BookingBuilderDialogProps> = ({
  open,
  bookingUrl,
  onClose,
  onOpen,
}) => {
  const [copied, setCopied] = useState(false);

 const handleCopy = async () => {
  try {
    await navigator.clipboard.writeText(bookingUrl);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  } catch {
    // Fallback for environments where clipboard API is unavailable
    const textArea = document.createElement("textarea");
    textArea.value = bookingUrl;
    textArea.style.position = "fixed";
    textArea.style.left = "-9999px";
    textArea.style.top = "-9999px";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    try {
      document.execCommand("copy");
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } finally {
      document.body.removeChild(textArea);
    }
  }
};
  return (
    <Dialog open={open} onClose={onClose} fullWidth maxWidth="sm">
      <DialogTitle sx={{ textAlign: "center", fontWeight: 700 }}>
         BookingBuilder®
      </DialogTitle>

      <DialogContent>
        <Typography variant="body1" sx={{ textAlign: "center", mb: 3 }}>
          Email has been processed successfully.
        </Typography>

        {/* URL Container */}
        <Box
          sx={{
            position: "relative",
            bgcolor: "#f8fafc",
            p: 2,
            pr: 6, // Space for the button
            borderRadius: "8px",
            border: "1px solid #e2e8f0",
          }}
        >
          <Typography
            variant="body2"
            sx={{
              color: "#64748b",
              wordBreak: "break-all",
              fontSize: "0.8rem",
              fontFamily: "monospace",
            }}
          >
            {bookingUrl}
          </Typography>

          {/* Copy Button */}
          <Tooltip title={copied ? "Copied!" : "Copy URL"}>
            <IconButton
              onClick={handleCopy}
              size="small"
              sx={{
                position: "absolute",
                top: 8,
                right: 8,
                color: copied ? "#22c55e" : "#94a3b8", 
                transition: "all 0.2s ease",
                "&:hover": {
                    bgcolor: copied ? "rgba(34, 197, 94, 0.04)" : "rgba(0,0,0,0.04)"
                }
              }}
            >
             
            </IconButton>
          </Tooltip>
        </Box>
      </DialogContent>

      <DialogActions sx={{ justifyContent: "space-evenly", gap: 1, pb: 3, px: 3 }}>
        <Button 
          variant="outlined" 
          onClick={onClose}
          sx={{ textTransform: 'none', borderRadius: '8px', flex: 1 }}
        >
          Cancel
        </Button>

        <Button
          variant="contained"
          onClick={onOpen}
          sx={{
            flex: 1,
            textTransform: 'none',
            borderRadius: '8px',
            backgroundColor: "#6aa2ba",
            "&:hover": { backgroundColor: "#5a8da3" },
          }}
        >
          Open
        </Button>
      </DialogActions>
    </Dialog>
  );
};

export default BookingBuilderDialog;