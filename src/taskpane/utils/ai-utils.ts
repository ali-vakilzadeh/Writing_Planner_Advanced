// AI utilities for plan generation and updates

// System prompt for generating a new plan
const NEW_PLAN_SYSTEM_PROMPT = `
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

// System prompt for updating an existing plan
const UPDATE_PLAN_SYSTEM_PROMPT = `
You are an expert writing assistant that helps users update their document plans.
The user will provide their current document plan and a description of how they want to update it.
Your task is to modify the plan according to their request.

IMPORTANT: Your response must be in valid JSON format with the same structure as the input:
{
  "planItems": [
    {
      "title": "Section Title",
      "level": 1,
      "comments": "Detailed instructions for writing this section"
    },
    // more sections...
  ]
}

Guidelines:
1. Preserve the existing structure where appropriate
2. You can add, remove, or modify sections as needed
3. Update the comments to reflect the user's new requirements
4. ONLY respond with the JSON, no other text
`

// Function to get the API key from localStorage
export const getApiKey = (): string => {
  if (typeof window !== "undefined") {
    return localStorage.getItem("openrouter_api_key") || ""
  }
  return ""
}

// Function to generate a new plan
export const generateNewPlan = async (prompt: string): Promise<any> => {
  const apiKey = getApiKey()
  if (!apiKey) {
    throw new Error("API key not set. Please configure your OpenRouter API key in settings.")
  }

  try {
    const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        "HTTP-Referer": "https://pmia.app",
        "X-Title": "Writing Planner Add-in",
      },
      body: JSON.stringify({
        model: "meta-llama/llama-4-maverick:free",
        messages: [
          { role: "system", content: NEW_PLAN_SYSTEM_PROMPT },
          { role: "user", content: prompt },
        ],
      }),
    })

    if (!response.ok) {
      const errorData = await response.json()
      throw new Error(`API Error: ${errorData.error?.message || response.statusText}`)
    }

    const data = await response.json()

    // Add proper null checks for the response structure
    if (!data || !data.choices || !Array.isArray(data.choices) || data.choices.length === 0) {
      console.error("Unexpected API response format:", data)
      throw new Error("Invalid response format from API. Please check your API key and model settings.")
    }

    const content = data.choices[0]?.message?.content

    if (!content) {
      throw new Error("No content returned from API")
    }

    // Try to parse the JSON response
    try {
      // Find JSON in the response (in case the AI included other text)
      const jsonMatch = content.match(/\{[\s\S]*\}/)
      const jsonString = jsonMatch ? jsonMatch[0] : content

      return JSON.parse(jsonString)
    } catch (parseError) {
      console.error("Error parsing JSON response:", parseError)

      // If JSON parsing fails, try to extract sections manually
      const sections = extractSectionsFromText(content)
      if (sections.length > 0) {
        return { planItems: sections }
      }

      throw new Error("Failed to parse AI response into a valid plan")
    }
  } catch (error) {
    console.error("Error generating new plan:", error)
    throw error
  }
}

// Function to update an existing plan
export const updateExistingPlan = async (currentPlan: any[], prompt: string): Promise<any> => {
  const apiKey = getApiKey()
  if (!apiKey) {
    throw new Error("API key not set. Please configure your OpenRouter API key in settings.")
  }

  try {
    // Create a simplified version of the plan to send to the AI
    const simplifiedPlan = currentPlan.map((item) => ({
      title: item.title,
      level: item.level,
      comments: item.comments || "",
    }))

    const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        "HTTP-Referer": "https://pmia.app",
        "X-Title": "Writing Planner Add-in",
      },
      body: JSON.stringify({
        model: "openai/gpt-3.5-turbo",
        messages: [
          { role: "system", content: UPDATE_PLAN_SYSTEM_PROMPT },
          {
            role: "user",
            content: `Here is my current plan:\n${JSON.stringify({ planItems: simplifiedPlan }, null, 2)}\n\nI want to update it as follows: ${prompt}`,
          },
        ],
      }),
    })

    if (!response.ok) {
      const errorData = await response.json()
      throw new Error(`API Error: ${errorData.error?.message || response.statusText}`)
    }

    const data = await response.json()

    // Add proper null checks for the response structure
    if (!data || !data.choices || !Array.isArray(data.choices) || data.choices.length === 0) {
      console.error("Unexpected API response format:", data)
      throw new Error("Invalid response format from API. Please check your API key and model settings.")
    }

    const content = data.choices[0]?.message?.content

    if (!content) {
      throw new Error("No content returned from API")
    }

    // Try to parse the JSON response
    try {
      // Find JSON in the response (in case the AI included other text)
      const jsonMatch = content.match(/\{[\s\S]*\}/)
      const jsonString = jsonMatch ? jsonMatch[0] : content

      return JSON.parse(jsonString)
    } catch (parseError) {
      console.error("Error parsing JSON response:", parseError)

      // If JSON parsing fails, try to extract sections manually
      const sections = extractSectionsFromText(content)
      if (sections.length > 0) {
        return { planItems: sections }
      }

      throw new Error("Failed to parse AI response into a valid plan")
    }
  } catch (error) {
    console.error("Error updating plan:", error)
    throw error
  }
}

// Helper function to extract sections from text if JSON parsing fails
const extractSectionsFromText = (text: string): any[] => {
  const sections = []
  const lines = text.split("\n")

  let currentSection: any = null

  for (const line of lines) {
    // Try to identify section titles (e.g., "1. Introduction" or "## Background")
    const titleMatch = line.match(/^(?:(?:\d+\.|#{1,2})\s+)?(.*?)(?::|-|$|$|$)/)

    if (titleMatch && titleMatch[1].trim().length > 0 && !line.includes(":") && line.length < 100) {
      // Determine level based on formatting
      const level = line.startsWith("#")
        ? line.startsWith("##")
          ? 2
          : 1
        : line.match(/^\d+\./)
          ? 1
          : currentSection
            ? currentSection.level + 1
            : 1

      // Save previous section if exists
      if (currentSection) {
        sections.push(currentSection)
      }

      // Start new section
      currentSection = {
        title: titleMatch[1].trim(),
        level: level,
        comments: "",
      }
    } else if (currentSection && line.trim().length > 0) {
      // Add line to current section comments
      currentSection.comments += line.trim() + " "
    }
  }

  // Add the last section
  if (currentSection) {
    sections.push(currentSection)
  }

  return sections
}
