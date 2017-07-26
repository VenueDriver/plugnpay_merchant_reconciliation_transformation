# PlugNPay merchant reconciliation transformation

Problem: Our PlugNPay merchant reconciliation reports list one row per transaction, but we need to allocate those transactions more deeply down to individual events.

Solution: A Ruby script for breaking out the transactions with additional details from a Venue Driver report for each ticket sale within a transaction.

## Running

### Step 1: (Optional) Install RVM

https://rvm.io/rvm/install

You might like to run Ruby in some other way.  Have fun with that.

### Step 2: Install Ruby

I used 2.4 and that's recorded in the .rvmrc file.  Any 2.x version will probably work.

### Step 3: Install Bundler

    gem install bundler

### Step 4: Install gem bundle

    bundle install

### Step 5: Run the script

    ruby transform.rb spreadsheet.xlsx
