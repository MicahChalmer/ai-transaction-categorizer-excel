/* global Excel console process */

// Import dotenv for environment variables
import dotenv from 'dotenv';

// Load environment variables from .env file
dotenv.config();

// Export environment variables for use in components
export const ENV = {
  GOOGLE_API_KEY: process.env.GOOGLE_API_KEY || '',
  OPENAI_API_KEY: process.env.OPENAI_API_KEY || ''
};
