// decodeURIComponent throws on a stray "%". Callers here decode text that came from outside
// (a caller's path, a Graph link) only to compare it, so an undecodable value is compared as
// written: it then fails the comparison it was going to fail anyway, instead of throwing.
export const safeDecode = (value: string): string => {
  try {
    return decodeURIComponent(value)
  } catch {
    return value
  }
}
