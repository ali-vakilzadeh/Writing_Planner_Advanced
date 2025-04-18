"use client"

import * as React from "react"
import {
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  PrimaryButton,
  DefaultButton,
  Stack,
  Label,
  Link,
  MessageBar,
  MessageBarType,
} from "@fluentui/react"
import { getApiKey, saveApiKey } from "../utils/api-utils"

interface SettingsDialogProps {
  isOpen: boolean
  onDismiss: () => void
}

export const SettingsDialog: React.FC<SettingsDialogProps> = ({ isOpen, onDismiss }) => {
  const [apiKey, setApiKey] = React.useState<string>("")
  const [saveMessage, setSaveMessage] = React.useState<string>("")
  const [messageType, setMessageType] = React.useState<MessageBarType>(MessageBarType.info)

  // Load the API key when the dialog opens
  React.useEffect(() => {
    if (isOpen) {
      setApiKey(getApiKey())
      setSaveMessage("")
    }
  }, [isOpen])

  const handleSave = () => {
    try {
      saveApiKey(apiKey.trim())
      setSaveMessage("API key saved successfully!")
      setMessageType(MessageBarType.success)

      // Close the dialog after a short delay
      setTimeout(() => {
        onDismiss()
      }, 1500)
    } catch (error) {
      setSaveMessage("Failed to save API key. Please try again.")
      setMessageType(MessageBarType.error)
    }
  }

  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: "API Settings",
        subText: "Configure your OpenRouter API key for AI functionality.",
      }}
      minWidth={400}
    >
      <Stack tokens={{ childrenGap: 15 }}>
        {saveMessage && <MessageBar messageBarType={messageType}>{saveMessage}</MessageBar>}

        <TextField
          label="OpenRouter API Key"
          value={apiKey}
          onChange={(_, newValue) => setApiKey(newValue || "")}
          placeholder="Enter your OpenRouter API key"
          type="password"
        />

        <Label>
          Don't have an API key? Get one at{" "}
          <Link href="https://openrouter.ai" target="_blank">
            openrouter.ai
          </Link>
        </Label>
      </Stack>

      <DialogFooter>
        <PrimaryButton onClick={handleSave} text="Save" />
        <DefaultButton onClick={onDismiss} text="Cancel" />
      </DialogFooter>
    </Dialog>
  )
}
