require 'win32ole'
require 'rspec-expectations'

class TestCompleteWorld
  def initialize(test_complete_path, project_name, script_unit)
    @project_name = project_name
    @script_unit = script_unit

    puts 'Connecting to TestComplete'
    @test_complete = WIN32OLE.connect('TestComplete.TestCompleteApplication')

    begin
      @test_complete.Integration
    rescue
      puts 'TestComplete does not appear to be running - starting instead'
      @test_complete = WIN32OLE.new('TestComplete.TestCompleteApplication')
    end

    test_complete_path = File.path(test_complete_path)

    puts "Connected to TestComplete - making visible and opening project #{test_complete_path}"

    @test_complete.Visible = true
    @test_complete.Integration.OpenProjectSuite(test_complete_path)

    @integration = @test_complete.Integration
  end

  def run_routine(name)
    puts "Running #{name} in project #{project_name}"
    begin
      run_with_delays do
        @integration.RunRoutine(@project_name, @script_unit, name)
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

    @test_complete.Integration.GetLastResultDescription.Status.should_not eq 2
  end

  def run_routine_ex(name, *args)
    puts "Running #{name} with arguments #{args} in project #{@project_name}"
    begin
      run_with_delays do
        @integration.RunRoutineEx(@project_name, @script_unit, name, args)
      end
    rescue
      raise "Call to #{name} with arguments #{args} failed"
    end
  end

  def call_script(name, *args)
    unless args.empty?
      run_routine_ex(name, args)
    else
      run_routine(name)
    end
  end
end
