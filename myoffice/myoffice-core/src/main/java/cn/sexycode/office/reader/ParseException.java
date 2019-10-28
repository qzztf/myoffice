package cn.sexycode.office.reader;

/**
 * Indicates that an expected getter or setter method could not be
 * found on a class.
 */
public class ParseException extends RuntimeException {
    /**
     * Constructs a PropertyNotFoundException given the specified message.
     *
     * @param message A message explaining the exception condition
     */
    public ParseException(String message) {
        super(message);
    }

    public ParseException(String message, Throwable cause) {
        super(message, cause);
    }
}
