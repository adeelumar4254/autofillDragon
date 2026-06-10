
import React, { createContext, useContext, useState } from "react";
import { Snackbar, Alert, Button } from "@mui/material";

type Severity = "success" | "error" | "warning" | "info";

interface SnackbarContextType {
  showSnackbar: (message: string, severity?: Severity, action?: () => void) => void;
}

const SnackbarContext = createContext<SnackbarContextType | undefined>(undefined);

export const SnackbarProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [open, setOpen] = useState(false);
  const [message, setMessage] = useState("");
  const [severity, setSeverity] = useState<Severity>("info");
  const [onAction, setOnAction] = useState<(() => void) | null>(null);

  const showSnackbar = (msg: string, sev: Severity = "info", action?: () => void) => {
    setMessage(msg);
    setSeverity(sev);
    setOnAction(() => action || null); 
    setOpen(true);
  };

  return (
     <SnackbarContext.Provider value={{ showSnackbar }}>
      {children}
      <Snackbar 
        open={open} 
        autoHideDuration={3000} 
        onClose={() => setOpen(false)}
        // ADD THIS LINE BELOW:
        anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
      >
        <Alert 
          onClose={() => setOpen(false)} 
          severity={severity} 
          variant="filled"
          sx={{ width: '100%' }} 
          action={
            onAction ? (
              <Button 
                color="inherit" 
                size="small" 
                onClick={() => {
                  onAction();
                  setOpen(false);
                }}
                sx={{ fontWeight: 800, textTransform: 'none' }}
              >
                REPORT
              </Button>
            ) : null
          }
        >
          {message}
        </Alert>
      </Snackbar>
    </SnackbarContext.Provider>
  );
};

export const useSnackbar = () => {
  const context = useContext(SnackbarContext);
  if (!context) throw new Error("useSnackbar must be used within SnackbarProvider");
  return context;
};