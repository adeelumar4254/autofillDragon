import React, { useEffect, useRef } from "react";
import { Box, Typography, LinearProgress } from "@mui/material";
import processingLoader from "../../../../assets/D_parsing-loader.webm";

interface ProcessingLoaderProps {
  open: boolean;
}

const ProcessingLoader = ({ open }: ProcessingLoaderProps) => {
  const videoRef = useRef<HTMLVideoElement>(null);

  useEffect(() => {
    if (open && videoRef.current) {
      videoRef.current.currentTime = 0;
      videoRef.current.play().catch(() => {});
    }
  }, [open]);

  if (!open) return null;

  return (
    <Box
      sx={{
        position: "fixed",
        inset: 0,
        zIndex: 99999,
        background: "#fff",
        display: "flex",
        flexDirection: "column",
        justifyContent: "center",
        alignItems: "center",
      }}
    >
      <video
        ref={videoRef}
        autoPlay
        muted
        loop
        playsInline
        style={{
          width: 210,
          height: 210,
          objectFit: "contain",
        }}
      >
        <source src={processingLoader} type="video/webm" />
      </video>

      <Typography
        variant="h6"
        sx={{
          mt: 2,
          fontWeight: 700,
          color: "#1e293b",
        }}
      >
        Processing Email
      </Typography>

      <Typography
        variant="body2"
        sx={{
          color: "#64748b",
          mt: 1,
          mb: 3,
        }}
      >
        Sending data to Booking Builder...
      </Typography>

      <LinearProgress
        sx={{
          width: 180,
          height: 6,
          borderRadius: 10,
          "& .MuiLinearProgress-bar": {
            backgroundColor: "#6aa2ba",
          },
        }}
      />
    </Box>
  );
};

export default ProcessingLoader;