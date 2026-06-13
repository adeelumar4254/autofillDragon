import React, { useState, useEffect, useCallback } from "react";
import {
    Box, Typography, Button, CircularProgress, Stack, Fade,
    Dialog, DialogTitle, DialogContent, DialogActions, Paper, Link,
    IconButton, MobileStepper
} from "@mui/material";
import {
    Email as EmailIcon,
    CheckCircle as CheckIcon,
    VpnKey as KeyIcon,
    Language as WebIcon,
    Phone as PhoneIcon,
    KeyboardArrowLeft,
    KeyboardArrowRight
} from "@mui/icons-material";
import { useNavigate } from "react-router-dom";
import { useSnackbar } from "../../context/SnackbarContext";
import Logo from "../../../../assets/BookingBuilder.png";
import { GetRegistrationEmail } from "../../services/Get_RegistrationEmail";
import { VerifyRegistrationNonce } from "../../services/Validate_RegisterCode";

const BRAND_DARK = "#1f5975";
const BRAND_LIGHT = "#6aa2ba";

function Register({ onConnected }: { onConnected: () => void }) {
    const navigate = useNavigate();
    const { showSnackbar } = useSnackbar();

    // Slider State
    const [activeStep, setActiveStep] = useState(0);
    const maxSteps = 2;

    const handleNext = () => setActiveStep((prev) => (prev + 1) % maxSteps);
    const handleBack = () => setActiveStep((prev) => (prev - 1 + maxSteps) % maxSteps);

    // Registration Logic States
    const [tokenInput, setTokenInput] = useState("");
    const [loading, setLoading] = useState(false);
    const [dismissedCodes, setDismissedCodes] = useState<string[]>([]);
    const [openRegDialog, setOpenRegDialog] = useState(false);
    const [openCodeDialog, setOpenCodeDialog] = useState(false);
    const [regLoading, setRegLoading] = useState(false);

    const userEmail = typeof Office !== "undefined"
        ? Office.context?.mailbox?.userProfile?.emailAddress || "Unknown Email"
        : "outlook.user@example.com";

    // --- Office Integration Logic ---
    const extractCodeFromBody = useCallback((text: string) => {
        if (!text) return;
        const regex = /Registration\s*code\s*:\s*([A-Za-z0-9_\-]+)/i;
        const match = text.match(regex);
        if (match && match[1]) {
            const detectedCode = match[1].trim();
            if (!dismissedCodes.includes(detectedCode) && tokenInput !== detectedCode) {
                setTokenInput(detectedCode);
                setOpenCodeDialog(true);
            }
        }
    }, [tokenInput, dismissedCodes]);

    const checkCurrentEmail = useCallback(() => {
        if (typeof Office !== "undefined" && Office.context?.mailbox?.item) {
            const item = Office.context.mailbox.item;
            item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    extractCodeFromBody(result.value);
                }
            });
        }
    }, [extractCodeFromBody]);

    useEffect(() => {
        const initTimer = setTimeout(checkCurrentEmail, 500);
        if (typeof Office !== "undefined" && Office.context?.mailbox) {
            const handler = () => checkCurrentEmail();
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, handler);
            return () => {
                clearTimeout(initTimer);
                Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged, handler);
            };
        }
        return () => clearTimeout(initTimer);
    }, [checkCurrentEmail]);

    const handleSendRegistrationEmail = async () => {
        setRegLoading(true);
        try {
            const response = await GetRegistrationEmail(userEmail);
            if (response.isSuccess) {
                showSnackbar("Registration email sent! check your inbox.", "success");
                setOpenRegDialog(false);
            } else {
                showSnackbar(response.errorMessage || "Failed to send email", "error");
            }
        } catch (err) {
            showSnackbar("An error occurred.", "error");
        } finally {
            setRegLoading(false);
        }
    };

    const handleRegisterDevice = async () => {
        setLoading(true);
        try {
            const response = await VerifyRegistrationNonce(tokenInput);
            if (response.isSuccess && response.data?.token) {
                localStorage.setItem("token", response.data.token);
                showSnackbar("Registered successfully!", "success");
                onConnected();
                navigate("/home");
            } else {
                showSnackbar(response.errorMessage || "Invalid code.", "error");
                setOpenCodeDialog(false);
            }
        } catch (err) {
            showSnackbar("Registration error.", "error");
        } finally {
            setLoading(false);
        }
    };

    return (
        <Box sx={{
            height: "100vh",
            display: "flex",
            flexDirection: "column",
            bgcolor: "#ffffff",
            overflowY: "Hidden",
            overflowX: "hidden"
        }}>
           
            <Box sx={{
                p: 3,
                paddingTop: "1px !important",
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center'
            }}>
                <Fade in={true} timeout={800}>
                    <Box sx={{ width: "100%", maxWidth: 350 }}>

                        {/* Logo Section */}
                        <Box sx={{ textAlign: "center", marginBottom: "10px" }}>
                            <Box sx={{
                                width: "250px",
                                height: "21vh",
                                mx: "auto",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center"
                            }}>
                                <img src={Logo} alt="Booking Builder Logo" style={{ width: '100%', objectFit: 'contain', marginBottom: "20px" }} />
                            </Box>
                            <Typography variant="h6" sx={{ fontWeight: 700, color: BRAND_DARK, fontSize: "0.90rem" }}>
                                Welcome to BookingBuilder® <br />
                                Autofill Dragon™
                            </Typography>
                        </Box>

                        {/* Information Card Slider */}
                        <Paper
                            elevation={0}
                            sx={{
                                bgcolor: "#fafafa",
                                minHeight: 220,
                                display: 'flex',
                                flexDirection: 'column',
                                paddingBottom: "0px !important",
                                marginTop: "30px"
                            }}
                        >
                            <Box sx={{ flexGrow: 1, paddingBottom: "0px", backgroundColor: "white !important" }} >
                                {activeStep === 0 ? (
                                    <Fade in={activeStep === 0}>
                                        <Box>
                                            <Typography variant="body2" sx={{ color: "#445", mb: 2, lineHeight: 1.5 }}>
                                                This add-in requires a subscription to <strong>Autofill Dragon™</strong>. If you do not have one, please contact us:
                                            </Typography>
                                            <Stack spacing={1}>
                                                <Link href="https://www.bookingbuilder.com" target="_blank" sx={{ display: 'flex', alignItems: 'center', gap: 1.5, fontSize: '0.85rem', textDecoration: 'none', color: BRAND_DARK, fontWeight: 700 }}>
                                                    <WebIcon fontSize="small" /> bookingbuilder.com
                                                </Link>
                                                <Link href="mailto:sales@bookingbuilder.com" sx={{ display: 'flex', alignItems: 'center', gap: 1.5, fontSize: '0.85rem', textDecoration: 'none', color: BRAND_DARK, fontWeight: 700 }}>
                                                    <EmailIcon fontSize="small" /> sales@bookingbuilder.com
                                                </Link>
                                                <Typography sx={{ display: 'flex', alignItems: 'center', gap: 1.5, fontSize: '0.85rem', color: BRAND_DARK, fontWeight: 700 }}>
                                                    <PhoneIcon fontSize="small" /> +1.845.234.4440
                                                </Typography>
                                            </Stack>
                                        </Box>
                                    </Fade>
                                ) : (
                                    <Fade in={activeStep === 1}>
                                        <Box>
                                            <Typography variant="body2" sx={{ color: "#445", mb: 2, lineHeight: 1.5 }}>
                                                Send booking emails to your GDS, ClientBase, or Tres with just a few clicks.
                                            </Typography>
                                            <Typography variant="body2" sx={{ color: BRAND_DARK, fontWeight: 700, mb: 1 }}>
                                                Click Register below to get started!
                                            </Typography>
                                            <Typography variant="caption" sx={{ color: "rgba(0, 0, 0, 1)", display: 'block', bgcolor: '#fff', p: 1, borderRadius: 1, border: '1px solid #eee' }}>
                                                <strong>Note:</strong> Ensure your admin has entered your email into your BookingBuilder® account.
                                            </Typography>
                                        </Box>
                                    </Fade>
                                )}
                            </Box>

                            <MobileStepper
                                variant="dots"
                                steps={maxSteps}
                                position="static"
                                activeStep={activeStep}
                                sx={{ bgcolor: 'white', flexGrow: 1, px: 0 }}
                                nextButton={
                                    <IconButton size="small" onClick={handleNext} sx={{ color: BRAND_DARK }}>
                                        <KeyboardArrowRight />
                                    </IconButton>
                                }
                                backButton={
                                    <IconButton size="small" onClick={handleBack} sx={{ color: BRAND_DARK }}>
                                        <KeyboardArrowLeft />
                                    </IconButton>
                                }
                            />
                        </Paper>

                        {/* Register Button - Fixed at Bottom */}
                        <Box sx={{
                            position: "fixed",
                            bottom: 0,
                            left: 0,
                            right: 0,
                            px: 3,
                            pb: 2,
                            pt: 1,
                            bgcolor: "#ffffff",
                            textAlign: "center"
                        }}>
                            <Button
                                variant="contained"
                                size="small"
                                fullWidth
                                disableElevation
                                onClick={() => setOpenRegDialog(true)}
                                sx={{
                                    py: 1.2,
                                    fontWeight: 800,
                                    borderRadius: "10px",
                                    bgcolor: BRAND_DARK,
                                    textTransform: "none",
                                    // fontSize: "0.85rem"
                                }}
                            >
                                Register
                            </Button>

                            <Typography
                                variant="caption"
                                sx={{
                                    display: "block",
                                    mt: 1.5,
                                    letterSpacing: 0.5,
                                    color: '#002d5b',
                                    fontWeight: 800,
                                    opacity: 0.5,
                                    fontSize: '0.65rem'
                                }}
                            >
                                © 2026 BookingBuilder Technologies • v1.0.0
                            </Typography>
                        </Box>

                    </Box>
                </Fade>
            </Box>




            {/* --- Dialogs logic  --- */}

            <Dialog
                open={openRegDialog}
                onClose={() => !regLoading && setOpenRegDialog(false)}
                disablePortal
                fullWidth={false}
                maxWidth={false}
                disableScrollLock
                slotProps={{
                    paper: {
                        sx: {
                            mx: 3,
                            width: "auto",
                            maxWidth: "280px",
                            borderRadius: "24px",
                            p: 1
                        }
                    }
                }}
            >
                <DialogTitle
                    sx={{
                        fontWeight: 800,
                        color: BRAND_DARK,
                        textAlign: "center"
                    }}
                >
                    Register Account
                </DialogTitle>

                <DialogContent>
                    <Typography
                        variant="body2"
                        sx={{
                            mb: 2,
                            textAlign: "center"
                        }}
                    >
                        We'll send a registration code to:
                    </Typography>

                    <Box
                        sx={{
                            p: 2,
                            bgcolor: "#f0f7f9",
                            borderRadius: "16px",
                            display: "flex",
                            alignItems: "center",
                            gap: 1.5,
                            overflow: "hidden"
                        }}
                    >
                        <Typography
                            variant="body2"
                            sx={{
                                fontWeight: 700,
                                color: BRAND_DARK,
                                wordBreak: "break-word",
                                overflowWrap: "anywhere",
                                fontSize: "0.85rem",
                                lineHeight: 1.4
                            }}
                        >
                            {userEmail}
                        </Typography>
                    </Box>
                </DialogContent>

                <DialogActions
                    sx={{
                        p: 2,
                        display: "flex",
                        justifyContent: "space-between",
                        gap: 1
                    }}
                >
                    <Button
                        onClick={() => setOpenRegDialog(false)}
                        disabled={regLoading}
                        sx={{
                            color: "text.secondary",
                            fontWeight: 700
                        }}
                    >
                        Cancel
                    </Button>

                    <Button
                        variant="contained"
                        onClick={handleSendRegistrationEmail}
                        disabled={regLoading}
                        sx={{
                            bgcolor: BRAND_DARK,
                            borderRadius: "10px",
                            minWidth: "120px"
                        }}
                    >
                        {regLoading ? (
                            <CircularProgress size={20} color="inherit" />
                        ) : (
                            "Send Code"
                        )}
                    </Button>
                </DialogActions>
            </Dialog>
            <Dialog
                open={openCodeDialog}
                onClose={() => !loading && setOpenCodeDialog(false)}
                disablePortal
                fullWidth={false}
                maxWidth={false}
                disableScrollLock
                slotProps={{
                    paper: {
                        sx: {
                            mx: 2,
                            width: "auto",
                            maxWidth: "280px",
                            borderRadius: "24px",
                            p: 1
                        }
                    }
                }}
            >
                <DialogTitle
                    sx={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        gap: 1,
                        fontWeight: 800,
                        color: BRAND_DARK
                    }}
                >
                    <CheckIcon color="success" />
                    Code Detected
                </DialogTitle>

                <DialogContent>
                    <Typography
                        variant="body2"
                        sx={{
                            mb: 2,
                            textAlign: "center"
                        }}
                    >
                        We found a registration code in this email.
                        Register now?
                    </Typography>

                    <Paper
                        elevation={0}
                        sx={{
                            p: 2,
                            bgcolor: "#f8fafc",
                            border: "2px dashed #6aa2ba",
                            textAlign: "center",
                            borderRadius: "12px",
                            overflow: "hidden"
                        }}
                    >
                        <Typography
                            variant="h6"
                            sx={{
                                fontWeight: 800,
                                color: BRAND_DARK,
                                wordBreak: "break-all",
                                letterSpacing: 2,
                            }}
                        >
                            {tokenInput}
                        </Typography>
                    </Paper>
                </DialogContent>

                <DialogActions
                    sx={{
                        p: 2,
                        flexDirection: "column",
                        gap: 1
                    }}
                >
                    <Button
                        fullWidth
                        variant="contained"
                        onClick={handleRegisterDevice}
                        disabled={loading}
                        sx={{
                            bgcolor: BRAND_DARK,
                            py: 1.4,
                            borderRadius: "12px",
                            fontWeight: 800,

                             textTransform: "none"
                        }}
                    >
                        {loading ? (
                            <CircularProgress size={22} color="inherit" />
                        ) : (
                            "Submit"
                        )}
                    </Button>

                    <Button
                        fullWidth
                        onClick={() => setOpenCodeDialog(false)}
                        sx={{
                            color: "text.secondary",
                            fontSize: "0.8rem"
                        }}
                    >
                        Not my code / Close
                    </Button>
                </DialogActions>
            </Dialog>



        </Box>
    );
}

export default Register;