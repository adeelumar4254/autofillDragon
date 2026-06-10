import React, { useState } from "react";
import { Box, Typography, TextField, Button, IconButton, Paper } from "@mui/material";
import { ArrowBack as BackIcon } from "@mui/icons-material";
import { useNavigate, useLocation } from "react-router-dom"; 
import { useSnackbar } from "../../context/SnackbarContext";

const ReportError = () => {
  const navigate = useNavigate();
  const location = useLocation(); 
  const { showSnackbar } = useSnackbar();
  
  // Retrieve the error passed from the Home screen
  const incomingError = location.state?.error || "No specific error data available.";

  const [description, setDescription] = useState("");
  const [submitting, setSubmitting] = useState(false);

  const handleSubmit = async () => {
    if (!description.trim()) {
      showSnackbar("Please describe the issue", "warning");
      return;
    }
    setSubmitting(true);
    
    
    console.log("Submitting Error Report:", { incomingError, userDescription: description });

    setTimeout(() => {
      showSnackbar("Error report submitted successfully", "success");
      setSubmitting(false);
      navigate("/home");
    }, 1500);
  };

  return (
    <Box sx={{ height: "100vh", bgcolor: "#ffffff", display: "flex", flexDirection: "column" }}>
      <Box sx={{ p: 2, display: "flex", alignItems: "center", borderBottom: "1px solid #f1f5f9" }}>
        <IconButton onClick={() => navigate("/home")} size="small" sx={{ mr: 1 }}>
          <BackIcon fontSize="small" />
        </IconButton>
        <Typography variant="h6" sx={{ fontWeight: 700, fontSize: "1rem" }}>Cancel</Typography>
      </Box>

      <Box sx={{ p: 2, flexGrow: 1, overflowY: 'auto' }}>
        {/* DISPLAY THE SYSTEM ERROR HERE */}
        <Typography variant="caption" sx={{ fontWeight: 700, color: "error.main", mb: 0.5, display: 'block' }}>
          {/* TECHNICAL ERROR DETECTED: */}
          DETECTED ERROR:

        </Typography>
        <Paper variant="outlined" sx={{ p: 1.5, mb: 3, bgcolor: '#fff5f5', borderColor: '#feb2b2', borderRadius: '8px' }}>
          <Typography variant="body2" sx={{ fontFamily: 'monospace', color: '#c53030', wordBreak: 'break-all' }}>
            {incomingError}
          </Typography>
        </Paper>

        <Typography variant="body2" sx={{ color: "#1f5975", mb: 1, fontWeight: 500 }}>
          Additional details (Optional):
           {/* Additional details (Optional): */}
        </Typography>
        <TextField
          fullWidth multiline rows={6}
          placeholder="Describe the issue you encountered"
          value={description}
          onChange={(e) => setDescription(e.target.value)}
          sx={{
            "& .MuiOutlinedInput-root": {
              bgcolor: "#f8fafc", borderRadius: "12px",
            }
          }}
        />
      </Box>

      <Box sx={{ p: 2, borderTop: "1px solid #f1f5f9" }}>
        <Button
          fullWidth size="small" variant="contained" disableElevation
          onClick={handleSubmit} disabled={submitting}
          sx={{ py: 1, bgcolor: "#1f5975", borderRadius: "10px", textTransform: "none", fontWeight: 700 }}
        >
          {submitting ? "Submitting..." : "Submit Report"}
        </Button>
      </Box>
    </Box>
  );
};

export default ReportError;