/**
 * Trail Journal -- Apps Script backend
 * File 2 of N: AI proxy -- faithful port of the live Supabase edge function `ai-proxy` (v15).
 *
 * Ported 1:1 from the actual deployed source (pulled directly from the Supabase project,
 * not reverse-engineered from the frontend). Same 14 prompt types, same prompt wording,
 * same model family. Only the transport changed: UrlFetchApp instead of Deno's fetch,
 * and the API key comes from Script Properties instead of a Supabase secret.
 *
 * Call this from the router (03_Router.gs) as: callAiProxy_(type, payload)
 * Returns the parsed JSON object (or {_raw, _parseError:true} if Claude's reply wasn't valid JSON).
 */

const CLAUDE_MODEL_ = 'claude-sonnet-4-6'; // matches the live ai-proxy function's MODEL constant

function callAiProxy_(type, payload) {
  const KEY = getProp_('ANTHROPIC_API_KEY');
  const prompt = buildAiPrompt_(type, payload);

  if (prompt === null) {
    return { error: 'Unknown type: ' + type };
  }

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': KEY,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL_,
      max_tokens: 1600,
      messages: [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Anthropic ' + response.getResponseCode() + ': ' + response.getContentText());
  }

  const ad = JSON.parse(response.getContentText());
  const raw = ((ad.content && ad.content[0] && ad.content[0].text) || '').trim().replace(/```json|```/g, '').trim();

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    parsed = { _raw: raw, _parseError: true };
  }

  // Mirrors the original: when a full "summary" comes back, also patch insight fields
  // onto the reflections row so the admin dashboard stats stay in sync.
  if (type === 'summary' && parsed.insightLevel && !parsed._parseError && payload.sessionId) {
    try {
      updateReflectionAiFields_(payload.sessionId, parsed);
    } catch (e) {
      // non-fatal, matches original try/catch-and-ignore behavior
    }
  }

  return parsed;
}

/**
 * Builds the exact prompt text for each type, word-for-word from the live function.
 * Returns null for an unrecognized type.
 */
function buildAiPrompt_(type, payload) {
  if (type === 'article') {
    const p = payload.part1;
    return 'You are a compassionate school counselor writing a short personalized reading for a middle school student at Camp Mountaineer. Trail/outdoor metaphors throughout.\n\n' +
      'Student: ' + payload.initial + '\n' +
      'What happened: ' + p.whatHappened + '\n' +
      'Why: ' + p.why + '\n' +
      'Thinking: ' + p.thinking + '\n' +
      'Emotions during (' + p.intensityDuring + '/5): ' + ((p.emotionsDuring || []).join(', ')) + '\n' +
      'Emotions now (' + p.intensityNow + '/5): ' + ((p.emotionsNow || []).join(', ')) + '\n' +
      'CPS before: ' + p.cpsUnsolvedProblem + '\n' +
      'My concern: ' + p.cpsMyConcern + '\n' +
      'Their concern: ' + p.cpsTheirConcern + '\n\n' +
      '3-4 warm paragraphs: (1) validate emotions without excusing, (2) brain science in accessible language, (3) connect to CPS insights while keeping the focus on the student\'s own choices, (4) forward-looking trail metaphor. ' +
      'OWNERSHIP CHECK: if the student\'s own words (why/thinking/CPS answers) blame someone else, minimize their own role, or make excuses ("they made me," "it wasn\'t my fault," "everyone else was doing it"), gently but clearly name that and redirect the reading toward what THEY chose and can control -- do not just validate the blame. Acknowledge the other person\'s role if relevant, but the throughline stays on the student\'s own actions. Do not shame them -- redirect, don\'t scold.\n\n' +
      'Return ONLY valid JSON: {"title":"trail-themed title","body":"paragraphs separated by \\n\\n"}';
  }

  if (type === 'hope_reflection') {
    const desc = { 1: 'cannot see a way forward', 2: "wants better but doesn't know how", 3: 'can see a path but needs help', 4: 'ready to do the work' };
    return 'School counselor responding to a student who rated hope ' + payload.hopeLevel + '/4 (' + (desc[payload.hopeLevel] || 'somewhere in the middle') + ').\n' +
      'What happened: ' + payload.whatHappened + '\n' +
      'Emotions: ' + ((payload.emotions || []).join(', ')) + '\n\n' +
      "2-3 warm paragraphs meeting them at their level, connecting hope to emotions, ending with one concrete action they can take right now. Be honest -- if they're at 1, honor that. " +
      "If \"what happened\" reads as blaming someone else rather than owning their part, gently redirect -- the concrete action should be something THEY control, not something that depends on someone else changing. " +
      'Return ONLY valid JSON: {"reflection":"paragraphs separated by \\n\\n"}';
  }

  if (type === 'cps_synthesis') {
    return 'School counselor helping a student understand both sides of a conflict.\n' +
      'Before: ' + payload.problem + '\n' +
      'My concern: ' + payload.myConcern + '\n' +
      'Their concern: ' + payload.theirConcern + '\n' +
      'What happened: ' + payload.whatHappened + '\n' +
      'Why: ' + payload.why + '\n\n' +
      '3-5 sentence paragraph acknowledging both sides, framing as two paths crossing -- but if "my concern" reads as blaming the other person entirely with no ownership of their own part, gently name that and bring the paragraph back to what the student themselves chose and can control. Forward-looking end. ' +
      'Return ONLY valid JSON: {"synthesis":"paragraph"}';
  }

  if (type === 'ripple_suggest') {
    return 'School counselor helping student think about who was affected.\n' +
      'What happened: ' + payload.whatHappened + '\n' +
      'Where: ' + (payload.where || 'not specified') + '\n' +
      'Involved: ' + (payload.whoInvolved || 'not specified') + '\n' +
      'Already named: ' + ((payload.named || []).join(', ')) + '\n\n' +
      '2-3 suggestions of people/groups NOT yet named. Name a TYPE, warm non-judgmental 1-2 sentence explanation. ' +
      'Return ONLY valid JSON: {"suggestions":[{"person":"type","insight":"how affected"}]}';
  }

  if (type === 'peak_guidance') {
    // Grounds the guidance in the actual EMS Mountaineer Peaks universal expectations deck
    // (11x17 EMS-Mountaineer-Peaks.pptx) -- each location's real "Be Present / Be
    // Personable / Be Productive" examples, not a generic definition. payload.where comes
    // from the Part 1 "Where did it happen?" dropdown, which now uses these exact location
    // names so this lookup always hits (falls back to a general description otherwise).
    const LOCATION_PEAKS_ = {
      'Classroom': 'Be Present: arrive on time, stay engaged, minimize distractions, follow routines. Be Personable: communicate respectfully with peers and staff, participate appropriately. Be Productive: focus on the task/lesson, follow instructions, complete work with integrity.',
      'Hallways': 'Be Present: eyes up, be aware of surroundings; be where you\'re supposed to be with a pass, on time. Be Personable: communicate kindly, polite words and body language, conversational volume, respect boundaries (hands/feet to self), help others. Be Productive: use the fastest route straight to your destination, stay with the flow of traffic.',
      'Cafeteria': 'Be Present: be there -- arrive on time, sit at a table, stay the whole period; be prepared with ID/lunch code. Be Personable: share the space, keep your area clean, honor personal space, communicate kindly (please/thank you/excuse me), engage with peers, be friendly. Be Productive: eat and prioritize your nutritional needs, clean up after yourself, use time wisely.',
      'Auditorium': 'Be Present: arrive on time, follow directions to your seating area, practice active listening (face the speaker, sit up, engage when invited). Be Personable: respond appropriately (right time/volume), be friendly. Be Productive: follow instructions, wait for dismissal directions, focus on learning.',
      'Bathroom': 'Be Present: have a pass, check in with your teacher first, use the nearest bathroom, keep trips short. Be Personable: respect others\' space (one person per stall), stay tech-free. Be Productive: use it as intended (not a hangout spot), keep it clean, return promptly (5 min or less), report issues to a trusted adult.',
      'Media Center': 'Be Present: have a pass, arrive ready to work with a clear purpose. Be Personable: share the space, be aware of others\' need for quiet, handle materials with care, communicate kindly in low voices. Be Productive: focus on your purpose (research/read/study/create), minimize distractions, show integrity, leave it better than you found it.',
      'Traveling To & From School': 'Be Present: be on time, know your route/stop, minimize distractions. Be Personable: communicate kindly with everyone around you, respect property and boundaries. Be Productive: plan ahead, follow laws and rules, pay attention and maintain safety, get where you\'re going.',
      'Digital Environment': 'Be Present: use your device as issued, keep your phone stored and off. Be Personable: communicate with intention (words carry weight and leave a record), honor others\' privacy before capturing or sharing images. Be Productive: use technology with purpose, show integrity (cite sources and AI tools used), own your account and protect your credentials, be a Mountaineer everywhere online, protect yourself and speak up if something feels wrong.',
    };
    const locNote = LOCATION_PEAKS_[payload.where] || null;

    return 'You are a compassionate school counselor at Ephrata Middle School helping a student see which Mountaineer Peaks universal expectation their incident relates to most, before they pick one themselves in the next step.\n\n' +
      'The three Peaks -- "Together We Climb":\n' +
      '- Be Present: showing up, on time, engaged, and aware of your surroundings.\n' +
      '- Be Personable: communicating with kindness and respect, sharing space well with others.\n' +
      '- Be Productive: using your time and the space you\'re in with purpose, following through, leaving things better than you found them.\n\n' +
      (locNote ? ('What "' + payload.where + '" specifically looks like at EMS:\n' + locNote + '\n\n') : '') +
      'What happened: ' + payload.whatHappened + '\n' +
      'Why: ' + payload.why + '\n' +
      'Thinking at the time: ' + payload.thinking + '\n\n' +
      'Pick the ONE Peak that best fits this specific situation. In 2-3 warm sentences, explain why -- reference their actual situation and, if given, the specific expectations for that location, not a generic definition. Speak directly to the student ("you"), second person. This is guidance to help them think it through, not a final decision -- they choose the Peak themselves next.\n\n' +
      'Return ONLY valid JSON: {"peak":"Be Present" or "Be Personable" or "Be Productive","whyThisFits":"2-3 sentences"}';
  }

  if (type === 'consequence_idea') {
    // Replaces the old static "Four Buckets" explainer tab (which showed the same 4
    // generic categories to every student regardless of situation) with a personalized
    // recommendation, generated right after Part 1 -- same timing as the 'article' call.
    // Grounded in the EMS Logical Consequences staff guide: Responsive Classroom's 3
    // foundational categories, with Curwin & Mendler's Discipline with Dignity framework
    // (Community, Teach Skills, Offer Choices, Transform, Inspire, Plan, Altruistic) as
    // optional supporting texture. This is guidance ahead of Part 2, not a final decision
    // -- the student still ranks and chooses their own consequence ideas there.
    return 'You are a compassionate school counselor helping a middle school student at Camp Mountaineer understand what kind of logical consequence tends to fit a situation like theirs, before they choose their own in the next step.\n\n' +
      'The three foundational categories (Responsive Classroom):\n' +
      '- "You Break It, You Fix It": something was broken, a mess was made, or harm was done to a person or relationship -- the student repairs it directly.\n' +
      '- "Loss of Privilege": behavior didn\'t meet the expectations tied to a privilege or material -- that privilege is removed, briefly and directly connected to the misuse.\n' +
      '- "Space & Time": the student needs a chance to calm down, regroup, and return -- not isolation or shame, called "Space and Time" at the middle school level.\n\n' +
      (payload.peak ? ('Mountaineer Peaks connection: this incident relates most to "' + payload.peak + '."\n\n') : '') +
      'What happened: ' + payload.whatHappened + '\n' +
      'Why: ' + payload.why + '\n' +
      'Thinking at the time: ' + payload.thinking + '\n' +
      'Emotions during: ' + ((payload.emotionsDuring || []).join(', ')) + '\n' +
      'Emotions now: ' + ((payload.emotionsNow || []).join(', ')) + '\n\n' +
      'Pick the ONE category that best fits this specific situation. In 2-3 warm sentences, explain why it fits -- reference their actual situation, not a generic definition. Then give ONE concrete example of what that could look like for them. Speak directly to the student ("you"), second person, warm but honest. This is meant to help them think it through, not tell them what will happen -- they choose the actual consequence themselves next.\n\n' +
      'Return ONLY valid JSON: {"category":"You Break It, You Fix It" or "Loss of Privilege" or "Space & Time","categoryIcon":"one emoji","whyThisFits":"2-3 sentences","example":"1-2 sentence concrete example"}';
  }

  if (type === 'consequence_nudge') {
    return 'School counselor reviewing consequence choices.\n' +
      'What happened: ' + payload.whatHappened + '\n' +
      'People hurt: ' + ((payload.hurtPeople || []).join(', ')) + '\n' +
      'Make right: ' + payload.makeRight + '\n' +
      'Ranked: ' + ((payload.ranked || []).join(' -> ')) + '\n\n' +
      'ONE warm sentence (max 30 words) affirming or gently nudging. ' +
      'Return ONLY valid JSON: {"nudge":"sentence"}';
  }

  if (type === 'creative_challenge') {
    return 'School counselor pushing creative thinking about consequence.\n' +
      'What happened: ' + payload.whatHappened + '\n' +
      'People hurt: ' + ((payload.hurtPeople || []).join(', ')) + '\n' +
      'Chosen: ' + ((payload.ranked || [])[0] || 'not selected') + '\n' +
      'Trait: ' + (payload.trait || 'not selected') + '\n\n' +
      '2-3 sentences affirming and pushing for more specific creative version. ' +
      'Return ONLY valid JSON: {"challenge":"2-3 sentence challenge"}';
  }

  if (type === 'skill_lesson') {
    const strats = {
      meta_moment: 'Meta-Moment: pause, ask What would my best self do?',
      granularity: 'Emotional Granularity: name emotions precisely',
      triggers: 'Trigger Mapping: identify specific triggers',
      breathing: 'Mindful Breathing 7-11',
      reframe: 'PAIR: Pause-Acknowledge-Investigate-Reframe',
      implementation: 'Implementation Intention: When X I will Y',
      attention: 'Attention Shifting',
      looping: 'Looping: Ask-Repeat-Check-Next Step',
      affective: 'Affective Statements',
      restorative_q: 'Restorative Questions',
      water: 'Water: metabolizes cortisol',
      look_up: 'Look Up: ceiling pause',
      forward_story: 'Forward Story: describe future self',
      pat: 'Pat Heart and Stomach',
    };
    return 'School counselor teaching strategy: ' + (strats[payload.strategy] || payload.strategy) + '\n' +
      'Domain: ' + payload.domain + ', Situation: ' + payload.whatHappened + ', Why: ' + payload.why +
      ', Emotions: ' + ((payload.emotions || []).join(', ')) + ', Trait: ' + payload.trait + '\n\n' +
      'Personalized micro-lesson with title, 1-sentence intro, 3-4 steps, 1-sentence application. ' +
      'Return ONLY valid JSON: {"title":"title","intro":"intro","steps":["s1","s2","s3"],"application":"application"}';
  }

  if (type === 'responder_note') {
    return 'Compassionate counselor validating a student affected by someone else\'s actions.\n' +
      'What happened: ' + payload.whatHappened + ', How affected: ' + payload.howAffected +
      ', Emotions (' + payload.intensity + '/5): ' + ((payload.emotions || []).join(', ')) +
      ', Needs: ' + payload.whatNeeds + ', Ready: ' + payload.readyToTalk + '\n\n' +
      '2-3 warm paragraphs validating and reassuring. ' +
      'Return ONLY valid JSON: {"title":"title","body":"paragraphs separated by \\n\\n"}';
  }

  if (type === 'parent_letter') {
    return 'AP writing parent communication. Professional, warm, solution-focused. No student names.\n' +
      'What happened: ' + payload.part1.whatHappened + ', Insight: ' + payload.insightLevel +
      ', Repair: ' + payload.repairPlan + ', Goal: ' + (payload.studentGoal || 'not set') +
      ', Trait: ' + payload.part3.trait + '\n\n' +
      '3 paragraphs: what occurred, what student committed to, invite to connect. ' +
      'Return ONLY valid JSON: {"subject":"subject","body":"paragraphs separated by \\n\\n"}';
  }

  if (type === 'pattern_analysis') {
    return 'AP reviewing anonymous restorative reflection data.\n' +
      'Data: ' + JSON.stringify(payload.submissions) + '\n' +
      'Date range: ' + payload.dateRange + '\n\n' +
      'Analyze triggers, emotions, insight levels, LRG traits, give 3 programming recommendations. ' +
      'Return ONLY valid JSON: {"summary":"overview","triggerPattern":"pattern","emotionPattern":"pattern","insightPattern":"suggestion","lrgPattern":"suggestion","recommendations":["r1","r2","r3"]}';
  }

  if (type === 'panel_inquiry_analysis') {
    const sr = payload.studentReflection || {};
    const tjNote = payload.tjData ? ('\n\nTrail Journal: insight ' + payload.tjData.insightLevel + ', trait ' + payload.tjData.trait + '.') : '';
    const notes = (payload.inquiryNotes || []).map(function (n) {
      return n.panelist + ': "' + n.question + '"\nStudent: "' + n.response + '"';
    }).join('\n\n');
    return 'You are an experienced school counselor analyzing what emerged from a Trailback Panel inquiry. Write for adults deliberating -- not for student consumption.\n\n' +
      'Incident: ' + payload.incident + '\n' +
      'Student initial: ' + payload.initial + tjNote + '\n\n' +
      'Student shared:\n' +
      '- What happened: ' + (sr.whatHappened || 'not captured') + '\n' +
      '- Who was affected: ' + (sr.whoAffected || 'not captured') + '\n' +
      '- Thought since: ' + (sr.thoughtSince || 'not captured') + '\n' +
      '- What they need: ' + (sr.whatNeeds || 'not captured') + '\n\n' +
      'Panel inquiry:\n' + notes + '\n\n' +
      'Write 2-3 paragraphs: (1) what the student genuinely understands about impact, (2) what they may still be avoiding or minimizing, (3) what they most need -- skill, support, connection, or structure. Direct and clinical. ' +
      'Return ONLY valid JSON: {"analysis":"paragraphs separated by \\n\\n"}';
  }

  if (type === 'panel_recommendation_draft') {
    const sr = payload.studentReflection || {};
    const d = payload.deliberation || {};
    const tjNote = payload.tjData ? ('\nTrail Journal: insight ' + payload.tjData.insightLevel + ', trait ' + payload.tjData.trait + ', goal: ' + (payload.tjData.goal || 'not set') + '.') : '';
    const impacts = (payload.impactStatements || []).filter(function (i) { return i.statement; })
      .map(function (i) { return (i.from || 'staff') + ': "' + i.statement + '"'; }).join('\n');
    return 'You are an experienced assistant principal drafting a Trailback Panel recommendation.\n\n' +
      'Incident: ' + payload.incident + '\n' +
      'Pattern: ' + (payload.pattern || 'none noted') + '\n' +
      'Student: ' + payload.initial + tjNote + '\n\n' +
      'Student shared: What happened: ' + (sr.whatHappened || '—') + ' | Who affected: ' + (sr.whoAffected || '—') +
      ' | Thinking since: ' + (sr.thoughtSince || '—') + ' | Needs: ' + (sr.whatNeeds || '—') + '\n\n' +
      'Impact statements:\n' + impacts + '\n\n' +
      'Deliberation: Harm: ' + (d.harm || '—') + ' | Student needs: ' + (d.studentNeeds || '—') + ' | Community needs: ' + (d.communityNeeds || '—') + '\n\n' +
      'Draft a Trailback Panel recommendation naming specific restorative actions tied to the harm, explaining why chosen, in a tone of restoration not punishment. Categories: Repair (fix/restore), Reconnect (dialogue/apology), Reflect (journal/project), Re-engage (school job/responsibility).\n\n' +
      'Return ONLY valid JSON: {"recommendation":"2-4 sentence narrative","rationale":"1-2 sentences explaining panel reasoning","suggestedCategories":["Category 1","Category 2"]}';
  }

  if (type === 'build-plan') {
    const p = payload;
    return 'You are a compassionate school counselor summarizing a student\'s growth plan at Camp Mountaineer. Write directly to the student in second person ("you"). Warm but honest.\n\n' +
      'Student: ' + p.initial + '\n' +
      (p.peak ? ('Mountaineer Peaks area: ' + p.peak + '\n') : '') +
      'What happened: ' + (p.whatHappened || 'not shared') + '\n' +
      'People hurt: ' + ((p.hurtPeople || []).join(', ') || 'not named') + '\n' +
      'How they\'ll make it right: ' + (p.makeRight || 'not stated') + '\n' +
      'What they\'d do differently: ' + (p.different || 'not stated') + '\n' +
      'Best idea for repair: ' + (p.bestIdea || 'not stated') + '\n' +
      'Plan: ' + (p.plan || 'not stated') + '\n' +
      'CPS unsolved problem: ' + (p.cpsUnsolvedProblem || 'not shared') + '\n' +
      'My concern: ' + (p.cpsMyConcern || 'not shared') + '\n' +
      'Their concern: ' + (p.cpsTheirConcern || 'not shared') + '\n' +
      'LRG Trait chosen: ' + (p.trait || 'not selected') + '\n' +
      'Specific areas: ' + ((p.traitSpecifics || []).join(', ') || 'none checked') + '\n' +
      'Why this trait: ' + (p.traitWhy || 'not explained') + '\n' +
      'Creative product: ' + (p.creativeFormat || 'not chosen') + '\n' +
      'Description: ' + (p.creativePlan || 'not described') + '\n' +
      'Community/service action: ' + (p.communityAction || 'not stated') + '\n\n' +
      'Write a 3-4 sentence summary that starts with "Your plan is to..." and weaves together: (1) the trait they\'re working on and why it matters for their situation, (2) what they\'re creating and how it connects to the harm, (3) how their repair commitment, creative product, and community/service action work together to both make things right AND leave their community better. End with one encouraging sentence. Keep it concise and direct -- this is a plan summary, not a pep talk.\n\n' +
      'Return ONLY valid JSON: {"planSummary":"the 3-4 sentence summary as a single string with no line breaks"}';
  }

  if (type === 'summary') {
    const p1 = payload.part1, p2 = payload.part2, p3 = payload.part3;
    const tier = payload.tier || 'full';
    const tierNote = {
      office: "This student's tier is OFFICE (shortest) -- only Part 1 plus a one-line repair note were ever collected by design. Part 2/Part 3 fields will be empty or near-empty -- that is expected, NOT avoidance. Do not penalize insight level for missing Part 2/3 content.",
      detention: "This student's tier is DETENTION -- Part 1 and Part 2 (repair plan) were collected; Part 3 (growth plan/trait) was never asked, by design. Do not penalize insight level for missing Part 3 content.",
      full: "This student's tier is FULL -- all three parts were collected.",
    }[tier];
    const cmNote = payload.cmReason ?
      ('CAMP MOUNTAINEER CONTEXT: Pathway: ' + payload.cmReason + ' | Initiated by: ' + (payload.cmInitiatedBy || 'not stated') +
       ' | Student-reported reason: ' + (payload.cmWhy || 'not stated') +
       // Student Checklist and Effort Agreement are both ISS-alternative-only pathway
       // requirements -- see CM_REASON_REQUIRES_EFFORT_ in index.html.
       (payload.cmReason === 'ISS' ? ' | Student Checklist confirmed: ' + (payload.cmChecklistAgreed ? 'Yes' : 'No') +
         ' | Effort Agreement confirmed: ' + (payload.cmEffortAgreed ? 'Yes' : 'No') : '') + '\n') : '';
    return 'Experienced middle school AP reviewing a completed Camp Mountaineer Trail Journal. Uses CPS, restorative practices, logical consequences.\n\n' +
      'STUDENT (initial: ' + payload.initial + '):\n' +
      (payload.location ? ('Location: ' + payload.location + ' | ') : '') +
      (payload.peak ? ('Mountaineer Peaks area: ' + payload.peak + '\n') : '\n') +
      cmNote +
      'PART 1: What happened: ' + p1.whatHappened + ' | Why: ' + p1.why + ' | Emotions during (' + p1.intensityDuring + '/5): ' +
      ((p1.emotionsDuring || []).join(',')) + ' | Emotions now (' + p1.intensityNow + '/5): ' + ((p1.emotionsNow || []).join(',')) +
      ' | CPS problem: ' + p1.cpsUnsolvedProblem + ' | My concern: ' + p1.cpsMyConcern + ' | Their concern: ' + p1.cpsTheirConcern + '\n' +
      'PART 2: People hurt: ' + ((p2.hurtPeople || []).join(',')) + ' | Make right: ' + p2.makeRight + ' | Consequences: ' +
      ((p2.selectedConsequences || []).join('-')) + ' | Plan: ' + p2.plan + ' | Different: ' + p2.different + '\n' +
      'PART 3: Trait: ' + p3.trait + ' | Specifics: ' + ((p3.traitSpecifics || []).join(',')) + ' | Why trait: ' + p3.traitWhy + ' | Creative: ' + p3.creativeFormat + ' | Community/service action: ' + (p3.communityAction || 'n/a') + '\n\n' +
      tierNote + '\n\n' +
      'INSIGHT LEVEL: HIGH = honest effort throughout + named someone hurt + any self-awareness + non-dismissive. Short honest answers count as High. MEDIUM = partial/thin engagement, OR effort is present but the account exclusively blames others with zero ownership of their own role. LOW = clearly avoidant/dismissive throughout only. Default strongly toward HIGH.\n\n' +
      'OWNERSHIP: if the student\'s account (why/thinking/makeRight/different) blames others without acknowledging their own part, do not let insightReason validate the blame -- name the pattern plainly for staff (e.g. "student attributes the incident primarily to [person/group] with limited ownership of their own actions") so it surfaces in the follow-up conversation.\n\n' +
      'Return this exact JSON:\n' +
      '{"summary":"2-3 sentences","insightLevel":"Low or Medium or High","insightReason":"1-2 sentences","studentGoal":"one goal second person",' +
      '"practicePlan":{"title":"Your Practice Plan","days":[{"day":"Today","task":"action"},{"day":"Day 2-3","task":"next"},{"day":"Day 4-5","task":"building"},{"day":"By Friday","task":"measurable"}]},' +
      '"recommendedConsequences":[{"type":"Primary","description":"logical","rationale":"why"},{"type":"Supporting","description":"skill building","rationale":"why"}],' +
      '"repairPlan":"2-3 steps","hopeNote":"one sentence for staff","restorativeQuestions":["q1","q2","q3","q4"],' +
      '"cpsScript":[{"speaker":"Adult","text":"empathy"},{"speaker":"Adult","text":"define problem"},{"speaker":"Adult","text":"invite solution"},{"speaker":"Adult","text":"what would be different?"},{"speaker":"Adult","text":"closing"}],' +
      '"followUpFlags":["flag"],' +
      '"studentReadingList":[{"title":"resource","topic":"concept","whyRelevant":"one sentence"},{"title":"...","topic":"...","whyRelevant":"..."},{"title":"...","topic":"...","whyRelevant":"..."}]}';
  }

  return null; // unknown type
}

/**
 * Mirrors the original edge function's post-summary PATCH back to the reflections row.
 */
function updateReflectionAiFields_(sessionId, parsed) {
  const sheet = getSheet_(TABS.REFLECTIONS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const sessionCol = headers.indexOf('session_id');
  const insightCol = headers.indexOf('insight_level');
  const summaryCol = headers.indexOf('ai_staff_summary');
  const readingCol = headers.indexOf('ai_reading_list');
  const flagsCol = headers.indexOf('has_flags');

  for (let i = 1; i < data.length; i++) {
    if (data[i][sessionCol] === sessionId) {
      const row = i + 1;
      sheet.getRange(row, insightCol + 1).setValue(parsed.insightLevel);
      sheet.getRange(row, summaryCol + 1).setValue(JSON.stringify(parsed));
      sheet.getRange(row, readingCol + 1).setValue(JSON.stringify(parsed.studentReadingList || []));
      sheet.getRange(row, flagsCol + 1).setValue((parsed.followUpFlags || []).length > 0);
      return;
    }
  }
}
