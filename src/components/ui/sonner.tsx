"use client"

import { useTheme } from "next-themes"
import { Toaster as Sonner, ToasterProps } from "sonner"

const Toaster = ({ ...props }: ToasterProps) => {
  const { theme = "system" } = useTheme()

  return (
    <Sonner
      theme={theme as ToasterProps["theme"]}
      className="toaster group"
      style={
        {
          "--normal-bg": "var(--popover)",
          "--normal-text": "var(--popover-foreground)",
          "--normal-border": "var(--border)",
          "--success-bg": "hsl(152, 69%, 97%)",
          "--success-text": "hsl(160, 84%, 24%)",
          "--success-border": "hsl(155, 55%, 80%)",
        } as React.CSSProperties
      }
      toastOptions={{
        classNames: {
          success: "!bg-emerald-50 !text-emerald-800 !border-emerald-200 dark:!bg-emerald-950/80 dark:!text-emerald-200 dark:!border-emerald-800/50",
          error: "!bg-red-50 !text-red-800 !border-red-200 dark:!bg-red-950/80 dark:!text-red-200 dark:!border-red-800/50",
        },
      }}
      {...props}
    />
  )
}

export { Toaster }
