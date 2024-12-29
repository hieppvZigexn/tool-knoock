if [ ! -f ./log/cron_log.log ]; then
  echo "Log file not found. Creating log.log..."
  touch ./log/cron_log.log
fi

bundle exec whenever --update-crontab && cron && tail -f ./log/cron_log.log
