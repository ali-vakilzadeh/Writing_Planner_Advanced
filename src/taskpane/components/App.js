"use client"

// Add this function to the App component to handle AI-generated plans
import * as React from "react"
import { AiPrompt } from "./AiPrompt"
import { useState } from "react"

// Declare Word variable
// /* global Word */ // Removed global declaration

const App = () => {
  const [nextId, setNextId] = useState(1)
  const [tocItems, setTocItems] = useState([])
  const [planningItems, setPlanningItems] = useState([])
  const [showTemplateButton, setShowTemplateButton] = useState(true)
  const [error, setError] = useState(null)
  const [isProcessingAi, setIsProcessingAi] = useState(false)

  const handleGeneratePlan = (planItems, isNewPlan) => {
    try {
      // Create new planning items from AI-generated plan
      const newNextId = nextId
      let newPlanningItems = []

      if (isNewPlan) {
        // For new plan, replace all existing items
        newPlanningItems = planItems.map((item, index) => {
          const id = newNextId + index
          return {
            id,
            title: item.title,
            level: item.level || 1,
            status: "empty",
            comments: item.comments || "",
            words: 0,
            paragraphs: 0,
            tables: 0,
            graphics: 0,
          }
        })

        // Create corresponding TOC items
        const newTocItems = newPlanningItems.map((item) => ({
          id: item.id,
          title: item.title,
          level: item.level,
          isDefault: true,
        }))

        // Update state
        setTocItems(newTocItems)
        setPlanningItems(newPlanningItems)
        setNextId(newNextId + newPlanningItems.length)
        setShowTemplateButton(false)
      } else {
        // For update plan, merge with existing items
        // First, create a map of existing items by title for easy lookup
        const existingItemsByTitle = {}
        planningItems.forEach((item) => {
          existingItemsByTitle[item.title.toLowerCase()] = item
        })

        // Process each item from the AI response
        newPlanningItems = planItems.map((item, index) => {
          // Check if we have an existing item with the same title
          const existingItem = existingItemsByTitle[item.title.toLowerCase()]

          if (existingItem) {
            // Update existing item
            return {
              ...existingItem,
              comments: item.comments || existingItem.comments,
            }
          } else {
            // Create new item
            const id = newNextId + index
            return {
              id,
              title: item.title,
              level: item.level || 1,
              status: "empty",
              comments: item.comments || "",
              words: 0,
              paragraphs: 0,
              tables: 0,
              graphics: 0,
            }
          }
        })

        // Update TOC items to match
        const newTocItems = newPlanningItems.map((item) => ({
          id: item.id,
          title: item.title,
          level: item.level,
          isDefault: true,
        }))

        // Update state
        setTocItems(newTocItems)
        setPlanningItems(newPlanningItems)
        setNextId(newNextId + newPlanningItems.filter((item) => !existingItemsByTitle[item.title.toLowerCase()]).length)
      }

      // Save to document properties
      setTimeout(() => saveToDocumentProperties(), 100)

      // Show success message
      setError(`Successfully ${isNewPlan ? "created" : "updated"} plan with ${newPlanningItems.length} sections`)
    } catch (error) {
      console.error("Error handling AI-generated plan:", error)
      setError("Failed to process AI-generated plan. Please try again.")
    }
  }

  // Add this function to generate content for a section
  const generateSectionContent = async (sectionId) => {
    const section = planningItems.find((item) => item.id === sectionId)
    if (!section) {
      setError("Section not found")
      return
    }

    // Check if comments are long enough
    if (!section.comments || section.comments.split(/\s+/).filter((word) => word.length > 0).length < 10) {
      setError("Section prompt must be at least 10 words long")
      return
    }

    setIsProcessingAi(true)
    setError(null)

    try {
      const apiKey = getApiKey()
      if (!apiKey) {
        throw new Error("API key not set. Please configure your OpenRouter API key in settings.")
      }

      // Prepare system prompt
      const systemPrompt = `You are an expert writing assistant. You are helping the user write content for a document section titled "${section.title}". The overall document is about writing planning. Please generate well-written, informative content based on the user's instructions.`

      // Make API call
      const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
          "HTTP-Referer": "https://writingplanner.app",
          "X-Title": "Writing Planner Add-in",
        },
        body: JSON.stringify({
          model: "openai/gpt-3.5-turbo",
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: section.comments },
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

      // Insert content into document
      try {
        // Check if Word API is available
        if (typeof Word === "undefined" || !Word || typeof Word.run !== "function") {
          console.error("Word API is not available")
          throw new Error("Word API is not available")
        }

        await Word.run(async (context) => {
          // Insert section title and content
          context.document.body.insertParagraph(section.title, "End").font.set({
            bold: true,
            size: 14,
          })
          context.document.body.insertParagraph(content, "End")
          context.document.body.insertParagraph("", "End")

          await context.sync()
        })

        // Update section status
        const updatedItems = planningItems.map((item) =>
          item.id === sectionId ? { ...item, status: "drafted" } : item,
        )
        setPlanningItems(updatedItems)
        setTimeout(() => saveToDocumentProperties(), 100)

        setError(`Content generated for "${section.title}" section`)
      } catch (error) {
        console.error("Error inserting content into document:", error)
        throw new Error("Failed to insert content into document")
      }
    } catch (error) {
      console.error("Error generating content:", error)
      setError(error instanceof Error ? error.message : "Failed to generate content")
    } finally {
      setIsProcessingAi(false)
    }
  }

  // Mock function to simulate saving to document properties
  // Replace this with your actual implementation
  const saveToDocumentProperties = () => {
    console.log("Saving to document properties (simulated)")
  }

  const getApiKey = () => {
    // Replace this with your actual implementation to retrieve the API key
    // For example, you might fetch it from local storage or a configuration file
    return localStorage.getItem("openRouterApiKey")
  }

  // Add this after the progress indicator and before the tabs
  // In the return statement of the App component, after the progress indicator and before the tabs section
  return (
    <div>
      {/* Your progress indicator here */}
      <AiPrompt
        existingPlan={planningItems.map((item) => ({
          id: item.id,
          title: item.title,
          level: item.level,
          comments: item.comments,
        }))}
        onGeneratePlan={handleGeneratePlan}
        onGenerateSectionContent={generateSectionContent}
        isProcessingAi={isProcessingAi}
      />
      {/* Your tabs section here */}
    </div>
  )
}

export default App
