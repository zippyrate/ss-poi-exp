package com.zippyrate.ss;

public class SSException extends Exception
{
	private static final long serialVersionUID = 1L;

	public SSException(String message, Throwable cause)
	{
		super(message, cause);
	}

	public SSException(String message)
	{
		super(message);
	}
}
