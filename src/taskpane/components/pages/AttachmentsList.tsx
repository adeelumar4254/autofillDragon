
import React from "react";
import {
  Box,
  Typography,
  Stack,
  Button,
  Paper,
  List,
  ListItem,
  Checkbox,
  Chip,
  Divider,
} from "@mui/material";
import { AttachmentInfo } from "../../types/email";
import { formatSize } from "../../utils/officeHelper";

const AttachmentsList = ({ attachments, selectedIds, onToggleOne, onSelectAll, glassStyle }: any) => {
  // Calculate if all valid attachments are currently selected
  const validAttachments = attachments.filter((a: any) => a.isValid);
  const isAllSelected = validAttachments.length > 0 && selectedIds.length === validAttachments.length;
 
 
  return (
    <Box sx={{ mt: 1 }}>
      <Stack direction="row" sx={{ px: 1, justifyContent: "space-between", alignItems: "center" }}>
        <Typography variant="overline" sx={{ fontWeight: 900, color: "#1f5975", letterSpacing: 2, opacity:1 }}>
          Assets 
          {/* ({attachments.length}) */}
        </Typography>
        {attachments.length > 0 && (
          <Button 
            size="small" 
            onClick={onSelectAll} 
            sx={{ 
              fontSize: '0.65rem', 
              fontWeight: 900, 
              textTransform: 'none', 
              color: isAllSelected ? 'error.main' : 'primary.dark' 
            }}
          >
            {isAllSelected ? "Unselect All" : "Select All"}
          </Button>
        )}
      </Stack>

      <Paper sx={{ ...glassStyle, mt: 1, maxHeight: 180, overflowY: 'auto' }}>
        <List dense disablePadding>
          {attachments.map((att: any, index: number) => (
            <React.Fragment key={att.id}>
              <ListItem 
                sx={{ 
                    py: 1, px: 2,
                    display: 'flex',
                    alignItems: 'flex-start',
                    transition: '0.2s',
                    bgcolor: selectedIds.includes(att.id) ? 'rgba(25, 118, 210, 0.03)' : 'transparent',
                    "&:hover": { bgcolor: 'rgba(0,0,0,0.02)' }
                }}
              >
                {/* 1. Checkbox */}
                <Checkbox 
                  size="small" 
                  checked={selectedIds.includes(att.id)} 
                  disabled={!att.isValid} 
                  onChange={() => onToggleOne(att.id)} 
                  sx={{ p: 0, mr: 1.5, mt: 0.2 }} 
                />

                {/* 2. Main Content Column */}
                <Box sx={{ flexGrow: 1, overflow: 'hidden' }}>
                  
                  {/* ROW 1: Name and Status */}
                  <Stack direction="row" spacing={1} sx={{ justifyContent: "space-between", alignItems: "center" }}>
                    <Typography 
                      variant="caption" 
                      sx={{ 
                        fontWeight: 800, 
                        color: att.isValid ? '#1e293b' : 'text.disabled', 
                        whiteSpace: 'nowrap', 
                        overflow: 'hidden', 
                        textOverflow: 'ellipsis',
                        maxWidth: '140px'
                      }}
                    >
                      {att.name}
                    </Typography>

                    <Chip
                      label={att.isValid ? "Eligible" : att.errorReason}
                      size="small"
                      sx={{ 
                        fontSize: '0.55rem', 
                        fontWeight: 900, 
                        height: 18,
                        flexShrink: 0, // Prevents chip from squishing
                        bgcolor: att.isValid ? 'rgba(34, 197, 94, 0.1)' : 'rgba(239, 68, 68, 0.1)',
                        color: att.isValid ? '#16a34a' : '#dc2626',
                        border: '1px solid rgba(0,0,0,0.03)'
                      }}
                    />
                  </Stack>

                  {/* ROW 2: Size Section */}
                  <Typography 
                    variant="caption" 
                    sx={{ 
                      fontSize: '0.6rem', 
                      opacity: 0.6, 
                      display: 'block',
                      mt: -0.2
                    }}
                  >
                    {formatSize(att.size)}
                  </Typography>

                </Box>
              </ListItem>
              {index < attachments.length - 1 && <Divider sx={{ opacity: 0.05 }} />}
            </React.Fragment>
          ))}
        </List>
      </Paper>
    </Box>
  );
};

export default AttachmentsList;