import express from 'express';
import OpenAI from "openai";
import cors from 'cors';

const app = express();
const openai = new OpenAI();

app.use(express.json());
// To allow a specific origin (your Office add-in development server):
const corsOptions = {
  origin: 'https://localhost:3000', // Ensure this matches the exact origin of your client
  optionsSuccessStatus: 200 // some legacy browsers (IE11, various SmartTVs) choke on 204
};


app.use(cors(corsOptions));

// Generic function to query GPT-4
async function queryGpt4(prompt) {
  const completion = await openai.chat.completions.create({
    messages: [{ role: "user", content: prompt }],
    model: "gpt-4",
  });
  return completion;
}

async function generateMetrics(email, sender, senderName, mainRecipient, mainRecipientName) {
  // Predict sport of email
  let sport_prompt = `Given the following email, which contains content based on the result of a sports event:\n${email}\n\nPlease predict what sport this email is talking about from the following choices: Soccer, Field Hockey, Girls Volleyball, Cross Country, Golf, Gymnastics, Baseball, Lacrosse, Outdoor Track, Tennis, Boys Volleyball, Basketball, Hockey, Swimming, Skiing, Wrestling, Indoor Track. In your response, list only the name of the sport you predict.`;
  let sport = await queryGpt4(sport_prompt);
  sport = sport.choices[0].message.content;

  let metrics = {
    "Soccer": "The final score, the score by halftime, all players who had a goal and/or assist in their totals, as well as both teams' goalkeepers and their total saves.",
    "Field Hockey": "The final score, the score by halftime, all players who had a goal and/or assist in their totals, as well as both teams' goalkeepers and their total saves.",
    "Cross Country": "The final team score, the location and distance of the race, and the full names of runners who finished in the Top 10, with each of their places and times.",
    "Girls Volleyball": "The final team score and the score by set, as well as the full name and stats for each player (Ace, Kills, Digs, Assists and Blocks)",
    "Golf": "The final team scores, as well as the full name and total strokes for each player.",
    "Gymnastics": "The final team scores, and the full names and scores of athletes who finish in top three of each event, as well as all-around.",
    "Baseball": "The final score, the score by innings, the full names of the winning in losing pitcher along with their innings pitched, earned runs, strikeouts, and walks and hits allowed, information for pitchers who earned saves, as well as the full name and stats of each player (at-bats, hits, extra-base hits, walks, and RBIs).",
    "Lacrosse": "The final score, the score by quarter, the full names of players who scoredgoals or made an assist, as well as the full names of goalies and the number of saves each one made.",
    "Outdoor Track": "The final score, the full names and time/distance of winners from each event, including relays, and notations of school records and qualifying times.",
    "Tennis": "The final score, the full name of the winners and losers from each match, as well as the score of each set (including tiebreakers) from each match.",
    "Boys Volleyball": "The final score, the score by set, and stats for each player",
    "Basketball": "The final score, the score by quarters, the team totals for field goals and foul shots made, the full names of every player who played, alongwith their total points with field goals made, foul shots made and total points, along with notation of 3-point field goals made.",
    "Hockey": "The final score, the score by periods, the full names of players who scored goals or made an assist, and the full names of goalies and the number of saves each made.",
    "Swimming": "The full name of the first place finisher, as well as their times for each event. If a relay was involved, the full names of each participant in the winning relay.",
    "Skiing": "The full name of the top five for each race, along with their times.",
    "Wrestling": "The full name of each wrestler and their opponent as well as the result of the match.",
    "Indoor Track": "The full name for the first-place finisher of each event. If a relay was involved, the full names of each participant in the winning relay."
  };

  let gpt_user_prompt = `I have an email below summarizing results from a high school ${sport.toLowerCase()} game:\n${email}`;
  gpt_user_prompt = gpt_user_prompt + `\nI would like to know if this email contains each of the following metrics: ${metrics[sport]}\n`;
  let test_prompt = gpt_user_prompt + "If ALL metrics are present, reply \"Yes\" and leave your response to that single word. Make a list that contains each of the metrics that are present, and a list that contains each of the metrics that are not present.";
  let evaluation = await queryGpt4(test_prompt);
  let email_response = "";
  let all_metrics = evaluation.choices[0].message.content;
  if (all_metrics == "Yes") {
    let confirm_prompt = `Pretend you are recording stats from a high school ${sport.toLowerCase()} event, and you are receiving these stats from a sender via email. Draft an email reply to the sender thanking them for providing them all of the metrics you need. Keep the email formal and within one or two sentences. The coach's name is ${mainRecipientName} and you should address them as Coach ${mainRecipientName.split(' ')[1]}. Your name which you should end with is ${senderName}. Use "Thank you" as the complimentary close. Do not include a subject.`;
    let confirm_evaluation = await queryGpt4(confirm_prompt);
    email_response = confirm_evaluation.choices[0].message.content;
    return email_response;
  }
  let not_included_prompt = `Given the following input text:\n${all_metrics}\nPlease identify all of the missing metrics in the input text and convert them into csv form and return the result. Note that the list may only contain one item.`;
  let evaluation2 = await queryGpt4(not_included_prompt);
  let not_included_metrics = evaluation2.choices[0].message.content;
  let email_prompt = `Pretend you are recording stats from a high school ${sport.toLowerCase()} event, and you are receiving these stats from a sender via email. Draft an email reply to the sender saying that you still need the following metrics from the event:\n${not_included_metrics}\nKeep the email formal and within a couple of sentences. The coach's name is ${mainRecipientName} and you should address them as Coach ${mainRecipientName.split(' ')[1]}. Your name which you should end with is ${senderName}. When you reference the list of missing metrics, please make that its own line. Use "Thank you" as the complimentary close. Do not include a subject. Here's a template for inspiration: 'Hi (coach’s name), \n Thanks for sending that information along. It looks like we’re missing (x, y, z) stat. Can you please send that additional information along as soon as possible?'`;
  let evaluation3 = await queryGpt4(email_prompt);
  email_response = evaluation3.choices[0].message.content;  
  return email_response;
}






// Endpoint to see what metrics an email does and does not contain
app.post('/metrics', async (req, res) => {
  const email = req.body.email;
  const sender = req.body.sender;
  const senderName = req.body.senderName;
  const mainRecipient = req.body.mainRecipient;
  const mainRecipientName = req.body.mainRecipientName;
  try {
    const metrics = await generateMetrics(email, sender, senderName, mainRecipient, mainRecipientName);
    res.json({ metrics });
  } catch (error) {
    console.error('Error in /metrics endpoint:', error);
    res.status(500).json({ error: 'An error occurred' });
  }
});


const port = process.env.PORT || 5000;
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});