// API utilities for OpenRouter integration

// Default API endpoint for OpenRouter
export const OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"

// Default system prompts
export const DEFAULT_PLANNER_SYSTEM_PROMPT = `
You are an expert writing assistant that helps users plan their documents.
The user will provide a description of what they want to write.
Your task is to create a structured document outline with sections and subsections.

IMPORTANT: Your response must be in valid JSON format with the following structure:
{
  "planItems": [
    {
      "title": "Section Title",
      "level": 1,
      "comments": "Detailed instructions for writing this section"
    },
    {
      "title": "Subsection Title",
      "level": 2,
      "comments": "Detailed instructions for writing this subsection"
    }
    // more sections...
  ]
}

Guidelines:
1. Level 1 is for main sections, level 2 is for subsections
2. Include 5-15 sections depending on the complexity of the topic
3. For each section, provide detailed comments (50-100 words) explaining what should be included
4. Follow standard academic or professional document structure when appropriate
5. ONLY respond with the JSON, no other text
`

export const DEFAULT_SECTION_SYSTEM_PROMPT = `
You are an expert writing assistant helping with a document.
The user needs content for a specific section.
Generate well-written, informative content based on the user's instructions.
Your response should be directly usable as the content for this section.
Focus on providing substantive, relevant information that fits the section's purpose.
`

// Function to get the API key from localStorage
export const getApiKey = (): string => {
  if (typeof window !== "undefined") {
    return localStorage.getItem("openRouterApiKey") || ""
  }
  return ""
}

// Function to save the API key to localStorage
export const saveApiKey = (apiKey: string): void => {
  if (typeof window !== "undefined") {
    localStorage.setItem("openRouterApiKey", apiKey)
  }
}

// Function to check if the API key is set
export const hasApiKey = (): boolean => {
  return !!getApiKey()
}

// Function to get the model from localStorage
export const getModel = (): string => {
  if (typeof window !== "undefined") {
    return localStorage.getItem("openrouter_model") || "openrouter/auto"
  }
  return "openrouter/auto"
}

// Function to get the planner system prompt from localStorage
export const getPlannerSystemPrompt = (): string => {
  if (typeof window !== "undefined") {
    return localStorage.getItem("planner_system_prompt") || DEFAULT_PLANNER_SYSTEM_PROMPT
  }
  return DEFAULT_PLANNER_SYSTEM_PROMPT
}

// Function to get the section system prompt from localStorage
export const getSectionSystemPrompt = (): string => {
  if (typeof window !== "undefined") {
    return localStorage.getItem("section_system_prompt") || DEFAULT_SECTION_SYSTEM_PROMPT
  }
  return DEFAULT_SECTION_SYSTEM_PROMPT
}

// Function to save system prompts to localStorage
export const saveSystemPrompts = (plannerPrompt: string, sectionPrompt: string): void => {
  if (typeof window !== "undefined") {
    localStorage.setItem("planner_system_prompt", plannerPrompt)
    localStorage.setItem("section_system_prompt", sectionPrompt)
  }
}

// Basic function to make API calls to OpenRouter
export const callOpenRouter = async (
  messages: Array<{ role: string; content: string }>,
  customModel?: string,
): Promise<any> => {
  const apiKey = getApiKey()
  const model = customModel || getModel()

  if (!apiKey) {
    throw new Error("API key not set. Please configure your OpenRouter API key in settings.")
  }

  try {
    const response = await fetch(OPENROUTER_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        "HTTP-Referer": "https://www.pmia.app", // Replace with your actual domain
        "X-Title": "Writer AI Add-in",
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
