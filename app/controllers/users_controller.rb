
require 'csv'
require 'caxlsx'

class UsersController < ApplicationController
  before_action :set_user, only: %i[ show edit update destroy ]

  # GET /users or /users.json
  def index
    @users = User.all
  end

  # GET /users/1 or /users/1.json
  def show
  end

  # GET /users/new
  def new
    @user = User.new
  end

  # GET /users/1/edit
  def edit
  end

  # POST /users or /users.json
  def create
    @user = User.new(user_params)

    respond_to do |format|
      if @user.save
        format.html { redirect_to user_url(@user), notice: "User was successfully created." }
        format.json { render :show, status: :created, location: @user }
      else
        format.html { render :new, status: :unprocessable_entity }
        format.json { render json: @user.errors, status: :unprocessable_entity }
      end
    end
  end

  # PATCH/PUT /users/1 or /users/1.json
  def update
    respond_to do |format|
      if @user.update(user_params)
        format.html { redirect_to user_url(@user), notice: "User was successfully updated." }
        format.json { render :show, status: :ok, location: @user }
      else
        format.html { render :edit, status: :unprocessable_entity }
        format.json { render json: @user.errors, status: :unprocessable_entity }
      end
    end
  end

  # DELETE /users/1 or /users/1.json
  def destroy
    @user.destroy!

    respond_to do |format|
      format.html { redirect_to users_url, notice: "User was successfully destroyed." }
      format.json { head :no_content }
    end
  end

  def export
    @users = User.all

    respond_to do |format|
      format.csv { send_data generate_csv(@users), filename: "users-#{Date.today}.csv" }
      format.xlsx { send_data generate_xlsx(@users), filename: "users-#{Date.today}.xlsx" }
    end
  end

  private
  def generate_csv(users)
    CSV.generate(headers: true) do |csv|
      csv << ["ID", "Name", "Email"]

      users.each do |user|
        csv << [user.id, user.name, user.email]
      end
    end
  end

  def generate_xlsx(users)
    p = Axlsx::Package.new
    wb = p.workbook
    wb.add_worksheet(name: "Users") do |sheet|
      sheet.add_row ["ID", "Name", "Email"]

      users.each do |user|
        sheet.add_row [user.id, user.name, user.email]
      end
    end
    puts "*************************testing"
    p.to_stream.read
  end

  # Use callbacks to share common setup or constraints between actions.
  def set_user
    @user = User.find(params[:id])
  end

  # Only allow a list of trusted parameters through.
  def user_params
    params.require(:user).permit(:name, :email)
  end
end
