import { toast } from 'sonner'

/** An error's own message when it has one, else `fallback`. */
export function errorMessage(error: unknown, fallback: string): string {
  return error instanceof Error && error.message !== '' ? error.message : fallback
}

/** A handler for a failed action that shows `title` and, when the error says more, why. */
export function reportFailure(title: string) {
  return (error: unknown): void => {
    toast.error(title, { description: error instanceof Error && error.message !== '' ? error.message : undefined })
  }
}
