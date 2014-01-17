require 'win32ole'
require 'rspec-expectations'

class TestCompleteWorld

  def initialize(test_complete_path, project_name)
    @project_name = project_name
    #@script_unit = ""

    begin
		puts 'Connecting to TestExecute'
	    @test_execute = WIN32OLE.connect('TestExecute.TestExecuteApplication')
		@test_execute.Integration
    rescue
		puts 'TestExecute does not appear to be running - starting instead'
		Thread.new {@test_execute = WIN32OLE.new('TestExecute.TestExecuteApplication')}
    end
	
	#dumb windows thing
    test_complete_path = test_complete_path.gsub! '/', '\\'

    puts "Connected to TestComplete - making visible and opening project #{test_complete_path}"

    #@test_execute.Visible = true
    @test_execute.Integration.OpenProjectSuite(test_complete_path)

    @integration = @test_execute.Integration
  end

  def run_routine(name,script_unit)
    puts "Running #{name} in project #{@project_name}"
    begin
      run_with_delays do
        @integration.RunRoutine(@project_name, script_unit, name)
      end
    rescue
      raise "Call to #{name} failed"
    end
  end

  def run_with_delays
    while @integration.IsRunning
      sleep 0.1
    end

    yield

    while @integration.IsRunning
      sleep 0.1
    end

    @test_execute.Integration.GetLastResultDescription.Status.should_not eq 2
  end

  def run_routine_ex(name, script_unit, *args)
    puts "Running #{name} with arguments #{args} in project #{@project_name}"
    begin
      run_with_delays do
        @integration.RunRoutineEx(@project_name, script_unit, name, args)
      end
    rescue
      raise "Call to #{name} with arguments #{args} failed"
    end
  end

  def call_script(name, script_unit, *args)
    unless args.empty?
      run_routine_ex(name,script_unit, args)
    else
      run_routine(name,script_unit)
    end
  end
end
