// API utilities for OpenRouter integration

// Default API endpoint for OpenRouter
export const OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"

// Function to get the API key from localStorage
export const getApiKey = (): string => {
  return localStorage.getItem("openrouter_api_key") || ""
}

// Function to save the API key to localStorage
export const saveApiKey = (apiKey: string): void => {
  localStorage.setItem("openrouter_api_key", apiKey)
}

// Function to check if the API key is set
export const hasApiKey = (): boolean => {
  return !!getApiKey()
}

// Basic function to make API calls to OpenRouter
export const callOpenRouter = async (
  messages: Array<{ role: string; content: string }>,
  model = "openai/gpt-3.5-turbo",
): Promise<any> => {
  const apiKey = getApiKey()

  if (!apiKey) {
    throw new Error("API key not set. Please configure your OpenRouter API key in settings.")
  }

  try {
    const response = await fetch(OPENROUTER_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        "HTTP-Referer": "https://writingplanner.app", // Replace with your actual domain
        "X-Title": "Writing Planner Add-in",
      },
      body: JSON.stringify({
        model: model,
        messages: messages,
      }),
    })

    if (!response.ok) {
      const errorData = await response.json()
      throw new Error(`API Error: ${errorData.error?.message || response.statusText}`)
    }

    return await response.json()
  } catch (error) {
    console.error("Error calling OpenRouter API:", error)
    throw error
  }
}
