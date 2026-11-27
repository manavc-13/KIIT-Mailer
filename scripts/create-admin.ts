// Create the first administrator, or promote/re-enable an existing user as admin.
//
//   npm run create-admin -- --email iqac.admin@kiit.ac.in --name "IQAC Administrator"
//
// Options:
//   --email <address>   account email (prompted if missing)
//   --name  <name>      display name (prompted if missing for a new user)
//   --password <pw>     set this password (avoid: it lands in shell history; omit to type it hidden)
//   --generate          generate a temporary password; the admin must change it at first sign-in
//   --help
//
// Idempotent: running it again for the same email promotes the user to admin and
// re-enables them; the password is only changed when you ask for it.
import { parseArgs } from 'node:util';
import readline from 'node:readline';
import { eq } from 'drizzle-orm';
import { db, schema } from '../server/db/client.js';
import { hashPassword, validatePasswordPolicy } from '../server/lib/passwords.js';
import { revokeUserSessions } from '../server/lib/session.js';
import { tempPassword } from '../server/lib/crypto.js';
import { audit } from '../server/lib/audit.js';
import { isValidEmail } from '../shared/text.js';

const HELP = `
Create or promote a KIIT Mailer administrator.

  npm run create-admin -- --email you@kiit.ac.in --name "Your Name" [--generate | --password <pw>]

Without --password or --generate you will be asked to type a password (hidden),
or to leave it empty to generate a temporary one.
`;

const { values: args } = parseArgs({
  options: {
    email: { type: 'string' },
    name: { type: 'string' },
    password: { type: 'string' },
    generate: { type: 'boolean', default: false },
    help: { type: 'boolean', short: 'h', default: false },
  },
  strict: true,
});

if (args.help) {
  console.log(HELP);
  process.exit(0);
}

const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: !!process.stdin.isTTY });
let muted = false;
// Hide typed characters while `muted` (only possible on an interactive terminal).
const rlAny = rl as unknown as { _writeToOutput?: (s: string) => void; output: NodeJS.WriteStream };
const originalWrite = rlAny._writeToOutput?.bind(rl);
rlAny._writeToOutput = (s: string) => {
  if (muted && !s.includes('\n') && !s.includes('\r')) return;
  if (originalWrite) originalWrite(s); else process.stdout.write(s);
};

function ask(question: string): Promise<string> {
  return new Promise(resolve => rl.question(question, a => resolve(a.trim())));
}

async function askHidden(question: string): Promise<string> {
  if (!process.stdin.isTTY) return ask(question);
  process.stdout.write(question);
  muted = true;
  const answer = await new Promise<string>(resolve => rl.question('', a => resolve(a)));
  muted = false;
  process.stdout.write('\n');
  return answer;
}

async function askYesNo(question: string, def = false): Promise<boolean> {
  const a = (await ask(`${question} ${def ? '[Y/n]' : '[y/N]'} `)).toLowerCase();
  if (!a) return def;
  return a === 'y' || a === 'yes';
}

function fail(msg: string): never {
  console.error(`\n  Error: ${msg}\n`);
  rl.close();
  process.exit(1);
}

/** Returns { password, temporary } chosen by flags or interactively. */
async function choosePassword(email: string): Promise<{ password: string; temporary: boolean }> {
  if (args.generate) return { password: tempPassword(), temporary: true };
  if (args.password !== undefined) {
    const err = validatePasswordPolicy(args.password, email);
    if (err) fail(err);
    return { password: args.password, temporary: false };
  }
  for (let attempt = 0; attempt < 3; attempt++) {
    const pw = await askHidden('Password (leave empty to generate a temporary one): ');
    if (!pw) return { password: tempPassword(), temporary: true };
    const err = validatePasswordPolicy(pw, email);
    if (err) { console.log(`  ${err}`); continue; }
    const again = await askHidden('Type the password again: ');
    if (again !== pw) { console.log('  The two passwords do not match.'); continue; }
    return { password: pw, temporary: false };
  }
  fail('No acceptable password entered.');
}

async function main() {
  console.log('\nKIIT Mailer: create or promote an administrator\n');

  let email = (args.email ?? '').trim().toLowerCase();
  if (!email) email = (await ask('Email address: ')).toLowerCase();
  if (!isValidEmail(email)) fail(`"${email}" is not a valid email address.`);

  const [existing] = await db.select().from(schema.users).where(eq(schema.users.email, email)).limit(1);

  if (existing) {
    const status = [existing.role === 'admin' ? 'administrator' : 'sender', existing.disabled ? 'disabled' : 'active'].join(', ');
    console.log(`A user with this email already exists: ${existing.name} (${status}).`);
    const name = args.name?.trim() || existing.name;

    let pw: { password: string; temporary: boolean } | null = null;
    if (args.generate || args.password !== undefined) pw = await choosePassword(email);
    else if (process.stdin.isTTY && await askYesNo('Set a new password for this account?', false)) pw = await choosePassword(email);

    const set: Partial<typeof schema.users.$inferInsert> = { role: 'admin', disabled: false, name, updatedAt: new Date() };
    if (pw) Object.assign(set, { passwordHash: await hashPassword(pw.password), mustChangePassword: pw.temporary });
    await db.update(schema.users).set(set).where(eq(schema.users.id, existing.id));
    if (pw) await revokeUserSessions(existing.id);
    await audit(null, 'user.update', {
      actorId: null, actorEmail: 'create-admin script', targetType: 'user', targetId: existing.id,
      meta: { email, changes: { role: { from: existing.role, to: 'admin' }, disabled: { from: existing.disabled, to: false } }, passwordReset: !!pw },
    });

    console.log(`\n  Done. ${name} <${email}> is an active administrator.`);
    if (pw?.temporary) printTemp(pw.password);
    else if (pw) console.log('  The new password is set. Existing sessions were signed out.');
    else console.log('  The password was not changed.');
    return;
  }

  let name = args.name?.trim() ?? '';
  if (!name) name = await ask('Full name: ');
  if (!name) fail('A name is required.');
  if (name.length > 120) fail('Keep the name under 120 characters.');

  const pw = await choosePassword(email);
  const [created] = await db.insert(schema.users).values({
    email, name, role: 'admin',
    passwordHash: await hashPassword(pw.password),
    mustChangePassword: pw.temporary,
  }).returning();
  await audit(null, 'user.create', {
    actorId: null, actorEmail: 'create-admin script', targetType: 'user', targetId: created.id,
    meta: { email, name, role: 'admin', via: 'cli' },
  });

  console.log(`\n  Done. Created administrator ${name} <${email}>.`);
  if (pw.temporary) printTemp(pw.password);
  else console.log('  Sign in with the password you just typed.');
}

function printTemp(pw: string) {
  console.log('\n  Temporary password (shown once; you will be asked to change it at first sign-in):');
  console.log(`\n      ${pw}\n`);
}

main()
  .then(() => { rl.close(); process.exit(0); })
  .catch((err) => { console.error('\n  Failed:', err instanceof Error ? err.message : err); rl.close(); process.exit(1); });
