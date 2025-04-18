// AI utilities for plan generation and updates

import { getApiKey, getModel, getPlannerSystemPrompt, getSectionSystemPrompt } from "./api-utils"

// Function to generate a new plan
export const generateNewPlan = async (prompt: string): Promise<any> => {
  const apiKey = getApiKey()
  const model = getModel()
  const systemPrompt = getPlannerSystemPrompt()

  if (!apiKey) {
    throw new Error("API key not set. Please configure your OpenRouter API key in settings.")
  }

  try {
    const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        "HTTP-Referer": "https://writingplanner.app",
        "X-Title": "Writer AI Add-in",
      },
      body: JSON.stringify({
        model: model,
        messages: [
          { role: "system", content: systemPrompt },
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
  const model = getModel()
  const systemPrompt = getPlannerSystemPrompt()

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
        "HTTP-Referer": "https://writingplanner.app",
        "X-Title": "Writer AI Add-in",
      },
      body: JSON.stringify({
        model: model,
        messages: [
          { role: "system", content: systemPrompt },
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

// Function to generate content for a specific section
export const generateSectionContent = async (
  sectionTitle: string,
  sectionComments: string,
  documentTitle = "Document",
): Promise<string> => {
  const apiKey = getApiKey()
  const model = getModel()
  const baseSystemPrompt = getSectionSystemPrompt()

  if (!apiKey) {
    throw new Error("API key not set. Please configure your OpenRouter API key in settings.")
  }

  try {
    // Create system prompt that focuses on the specific section
    const systemPrompt = `${baseSystemPrompt}\nThe document is titled "${documentTitle}" and the section is "${sectionTitle}".`

    const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        "HTTP-Referer": "https://writingplanner.app",
        "X-Title": "Writer AI Add-in",
      },
      body: JSON.stringify({
        model: model,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: sectionComments },
        ],
      }),
    })

    if (!response.ok) {
      const errorData = await response.json()
      throw new Error(`API Error: ${errorData.error?.message || response.statusText}`)
    }

    const data = await response.json()
    const content = data.choices[0]?.message?.content

    if (!content) {
      throw new Error("No content returned from API")
    }

    return content
  } catch (error) {
    console.error("Error generating section content:", error)
    throw error
  }
}
