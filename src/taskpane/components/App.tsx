import React, { useEffect, useState, useCallback } from "react";
import {
  Container,
  Typography,
  Card,
  CardContent,
  List,
  ListItem,
  ListItemText,
  Button,
  Divider,
  Box,
  CircularProgress,
  IconButton,
  Tooltip,
} from "@mui/material";
import { ArrowDownload16Filled, ArrowDownload16Regular, Link12Filled } from "@fluentui/react-icons";
// import DownloadIcon from '@mui/icons-material/Download';
// import LinkOutlinedIcon from '@mui/icons-material/LinkOutlined';

// --- Utility Functions (Strict Logic) ---

const downloadableExtensions = [
  ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rar", ".csv", ".txt",
];

// Target domain for strict filtering
const TARGET_DOMAIN = "cdn.azams0.fds.api.mi-img.com";

const isDownloadable = (url: string) => {
    const lowerUrl = url.toLowerCase();

    // 1. Strict Check: Agar URL target domain se nahi hai, toh reject kar dein
    if (!lowerUrl.includes(TARGET_DOMAIN)) {
        // Agar aap chahte hain ki TARGET_DOMAIN ke alawa koi aur standard file (.pdf) bhi na aaye, 
        // toh sirf TARGET_DOMAIN check kaafi hai. Agar koi bhi standard file bhi chahiye, 
        // toh niche wala logic use karein.
        
        // **Current Goal:** Sirf specific domain ke links ya standard files.
        
        // Standard file extension check (Query param safe)
        const cleanUrl = url.split(/[?#]/)[0].toLowerCase();
        return downloadableExtensions.some((ext) => cleanUrl.endsWith(ext));
    }
    
    // Target domain ko hamesha allow karo
    return true;
};

const extractLinks = (html: string): string[] => {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const anchors = Array.from(doc.querySelectorAll("a"));
  
  let links = anchors.map((a) => a.href);

  // Filtering: sirf http/https links jo downloadable hain
  return links
    .filter((href) => href.startsWith("http") && isDownloadable(href));
};

const processItemBody = (item: Office.MessageRead): Promise<string[]> => {
    return new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(extractLinks(result.value));
        } else {
          console.error("Failed to get body:", result.error);
          reject(result.error);
        }
      });
    });
};

// --- React Component ---

const App: React.FC = () => {
  const [allLinks, setAllLinks] = useState<string[]>([]);
  const [loading, setLoading] = useState<boolean>(true);

  // Sequential Item Handler
  const loadAndUnloadSingleItem = (itemId: string): Promise<string[]> => {
    return new Promise((resolve) => {
      
      Office.context.mailbox.loadItemByIdAsync(itemId, (loadResult) => {
        if (loadResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(
            `Failed to load item ${itemId} (Code: ${loadResult.error.code}): ${loadResult.error.message}`
          );
          return resolve([]);
        }

        const loadedItem:any = loadResult.value as unknown as Office.MessageRead;

        processItemBody(loadedItem)
          .then(resolve) 
          .catch((e) => {
            console.error(`Body processing failed for ${itemId}`, e);
            resolve([]); 
          })
          .finally(() => {
            loadedItem.unloadAsync(() => {
              // Unload complete
            });
          });
      });
    });
  };

  // Main processing loop (Duplicates allowed)
  const loadAndProcessSelectedItems = useCallback(() => {
    setLoading(true);
    setAllLinks([]);

    Office.context.mailbox.getSelectedItemsAsync(async (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        setLoading(false);
        return;
      }

      const selectedItems = asyncResult.value as Office.SelectedItemDetails[];
      if (selectedItems.length === 0) {
        setLoading(false);
        return;
      }
      
      const itemIds = selectedItems.map(item => item.itemId);
      let extractedLinks: string[] = [];

      for (const itemId of itemIds) {
        const linksFromItem = await loadAndUnloadSingleItem(itemId); 
        
        // *** DUPLICATE LINKS ALLOWED: Directly push all links ***
        linksFromItem.forEach(link => {
            extractedLinks.push(link);
        });
      }

      setAllLinks(extractedLinks);
      setLoading(false);
    });
  }, []); 

  // Initialization and Event Listener Setup (Same as before)
  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Outlook) {
        if (Office.context.requirements.isSetSupported("Mailbox", "1.15")) {
          Office.context.mailbox.addHandlerAsync(
            Office.EventType.SelectedItemsChanged,
            loadAndProcessSelectedItems,
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                loadAndProcessSelectedItems();
              }
            }
          );
        } else {
          setLoading(false);
        }
      }
    });
  }, [loadAndProcessSelectedItems]);

  const downloadAllLinks = () => {
    allLinks.forEach((link) => {
      const a = document.createElement("a");
      a.href = link;
      a.target = "_blank";
      a.click();
    });
  };

  // --- Attractive UI Rendering ---
  return (
    <Container maxWidth="sm" sx={{ p: 0, pt: 1 }}>
      <Typography
        variant="h5"
        align="center"
        sx={{ 
            fontWeight: 'bold', 
            color: '#106ebe', 
            mb: 2, 
            p: 1,
            borderRadius: 1,
            backgroundColor: '#f5f8ff'
        }}
      >
        Multi-Email Downloader
      </Typography>

      <Card
        sx={{ 
          maxHeight: 350, 
          overflowY: "auto", 
          mb: 3, 
          boxShadow: 3, 
          borderRadius: 2,
          border: '1px solid #e0e0e0'
        }}
      >
        <CardContent sx={{ p: 0 }}>
          {loading ? (
            <Box display="flex" flexDirection="column" alignItems="center" p={4}>
              <CircularProgress size={30} sx={{ mb: 1, color: '#106ebe' }} />
              <Typography variant="body2" color="text.secondary">
                Analyzing {allLinks.length > 0 ? 'new' : 'selected'} messages...
              </Typography>
            </Box>
          ) : allLinks.length === 0 ? (
            <Typography variant="body1" color="text.secondary" align="center" p={3}>
              No target downloadable files found.
            </Typography>
          ) : (
            <List disablePadding>
              {allLinks.map((link, idx) => (
                <React.Fragment key={idx}>
                  <ListItem
                    // key must be unique per item, using index is fine since we allow duplicates
                    key={idx} 
                    secondaryAction={
                        <Tooltip title="Download this file">
                            <IconButton 
                                edge="end" 
                                aria-label="download" 
                                href={link}
                                target="_blank"
                                size="small"
                                sx={{ color: '#106ebe' }}
                            >
                                {/* <ArrowDownload16Filled fontSize="small"/> */}
                            </IconButton>
                        </Tooltip>
                    }
                    sx={{ 
                      px: 2, 
                      py: 0.5,
                      backgroundColor: idx % 2 === 0 ? '#fcfcfc' : 'white', 
                      "&:hover": { backgroundColor: "#e3f2fd" },
                    }}
                  >
                    <ListItemText
                      primary={
                        <Box display="flex" alignItems="center">
                            <Link12Filled style={{ fontSize: '0.9rem', marginLeft: 1, color: '#0d47a1' }} />
                            {link}
                        </Box>
                      }
                      primaryTypographyProps={{
                        fontSize: "0.75rem",
                        color: "#555",
                        noWrap: true,
                        sx: {
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          fontFamily: "monospace",
                        },
                      }}
                    />
                  </ListItem>
                  <Divider component="li" light />
                </React.Fragment>
              ))}
            </List>
          )}
        </CardContent>
      </Card>

      {allLinks.length > 0 && (
        <Box textAlign="center">
          <Button
            variant="contained"
            size="large"
            onClick={downloadAllLinks}
            disabled={loading}
            startIcon={<ArrowDownload16Regular />}
            sx={{
              fontWeight: "bold",
              px: 5,
              py: 1.2,
              fontSize: '1rem',
              backgroundColor: "#2e7d32", // Green button for download
              "&:hover": { backgroundColor: "#1b5e20" },
              boxShadow: 4,
            }}
          >
            Download All Files ({allLinks.length})
          </Button>
        </Box>
      )}
    </Container>
  );
};

export default App;